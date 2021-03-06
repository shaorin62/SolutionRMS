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
Public Class ccPDCOPONO
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccPDCOPONO"                  '자신의 클래스명
    Private mobjcePD_PONO As ePDCO.cePD_PONO             '사용할 Entity 변수 선언
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password"
#End Region

#Region "GROUP BLOCK : Function Section"

   

    ' =============== Project Code 조회
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strFROM As String, _
                              ByVal strTO As String, _
                              ByVal strPONAME As String, _
                              ByVal strPONO As String, _
                              ByVal strCLIENTNAME As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCHOICE As String, _
                              ByVal cmbPOPUPTYPE As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)
        Dim Con1, Con2, Con3, Con4, Con5

        strCols = " PROJECTNO,"
        strCols = strCols & " PROJECTNM,"
        strCols = strCols & " CLIENTCODE,"
        strCols = strCols & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
        strCols = strCols & " SUBSEQ,"
        strCols = strCols & " DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME,"
        If strCHOICE = "ST" Then
            strCols = strCols & " CASE ISNULL(GROUPGBN,'') WHEN '2' THEN '그룹' ELSE '비그룹' END AS GROUPGBN,"
        Else
            strCols = strCols & " GROUPGBN,"
        End If
        strCols = strCols & " CREDAY,"
        strCols = strCols & " CPDEPTCD,"
        strCols = strCols & " DBO.SC_DEPT_NAME_FUN(CPDEPTCD) CPDEPTNAME,"
        strCols = strCols & " CPEMPNO,"
        strCols = strCols & " DBO.SC_EMPNAME_FUN(CPEMPNO) CPEMPNAME,"
        strCols = strCols & " MEMO,"
        strCols = strCols & " TIMCODE,"
        strCols = strCols & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) CLIENTTEAMNAME"

        If strFROM <> "" And strTO <> "" Then
            Con1 = String.Format(" AND (CREDAY BETWEEN '{0}' AND  '{1}')", strFROM, strTO)
        End If
        If strFROM <> "" And strTO = "" Then
            Con1 = String.Format(" AND (CREDAY >= '{0}')", strFROM)
        End If
        If strFROM = "" And strTO <> "" Then
            Con1 = String.Format(" AND (CREDAY <= '{0}')", strTO)
        End If


        'If strPONAME <> "" Then Con2 = String.Format(" AND (PROJECTNM like '%{0}%')", strPONAME)
        'If strPONO <> "" Then Con3 = String.Format(" AND (PROJECTNO = '{0}')", strPONO)
        If cmbPOPUPTYPE = "1" Then
            If strPONO <> "" Then Con2 = String.Format(" AND (PROJECTNO = '{0}')", strPONO)

            strPONAME = Replace(strPONAME, "'", "''")
            If strPONAME <> "" Then Con3 = String.Format(" AND (PROJECTNM LIKE '%{0}%')", strPONAME)

        Else
            If strPONO <> "" Then Con2 = String.Format(" AND (PROJECTNO = DBO.PD_GET_JOBNO_PROJECTNO_FUN('{0}'))", strPONO)

            strPONAME = Replace(strPONAME, "'", "''")
            If strPONAME <> "" Then Con3 = String.Format(" AND (PROJECTNO = DBO.PD_GET_JOBNAME_PROJECTNO_FUN('{0}'))", strPONAME)
        End If


        If strCLIENTNAME <> "" Then Con4 = String.Format(" AND (LTRIM(DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)) like '%{0}%')", strCLIENTNAME)
        If strCLIENTCODE <> "" Then Con5 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)
        strFormat = "SELECT {0} FROM PD_PONO WHERE 1=1 {1} ORDER BY PROJECTNO DESC"

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

    ' =============== Project Code 조회
    Public Function SelectRtn_PROJECTLIST(ByVal strInfoXML As String, _
                                          ByRef intRowCnt As Integer, _
                                          ByRef intColCnt As Integer, _
                                          ByVal strPROJECTLIST As String) As Object

        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체

            strFormat = " SELECT "
            strFormat = strFormat & " PROJECTNO,"
            strFormat = strFormat & " PROJECTNM,"
            strFormat = strFormat & " CREDAY,"
            strFormat = strFormat & " '' ENDDAY,"
            strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTCODE, "
            strFormat = strFormat & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMCODE,"
            strFormat = strFormat & " DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ)SUBSEQ,"
            strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(CPDEPTCD) CPDEPTCD,"
            strFormat = strFormat & " 'MC' + SUBSTRING(CPEMPNO,4, LEN(CPEMPNO))"
            strFormat = strFormat & " FROM PD_PONO"
            strFormat = strFormat & " WHERE PROJECTNO IN({0})"

            strSQL = String.Format(strFormat, strPROJECTLIST)
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_PROJECTLIST")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' ============== ProcessRtn PROJECT 코드 등록
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal strMasterXML As String, _
                               ByVal strSEQFLAG As String) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer '결과값 변수
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보를 가지고 있는 Config 개체
            Try
                'XML Element 변수 선언 (strMasterXML을 변환)
                Dim xmlRoot As XmlElement
                xmlRoot = XMLGetRoot(strMasterXML) 'XML 데이터

                'DB접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master 데이터 처리
                intRtn = ProcessRtn_Seq(xmlRoot, strSEQFLAG)

                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                'Resource해제
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' ============== ProcessRtn JOB 코드 등록
    Public Function ProcessRtnSheet_Insert(ByVal strInfoXML As String, _
                                           ByVal vntData As Object, _
                                           ByRef strPROJECTNO As String, _
                                           ByRef strPROJECTLIST As String) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer '결과값 변수
        Dim i, intColCnt, intRows, intRow As Integer  '루프, 컬럼Cnt, 로우Cnt 변수
        Dim strNEWPROJECTNO
        Dim strYEAR
        Dim strCREDAY, strGROUPGBN
        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_PONO = New cePD_PONO(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        If GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strCREDAY = GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)

                        If GetElement(vntData, "GROUPGBN", intColCnt, i, OPTIONAL_STR) = "그룹" Then
                            strGROUPGBN = "2"
                        Else
                            strGROUPGBN = "1"
                        End If

                        '실제저장부분
                        If GetElement(vntData, "PROJECTNO", intColCnt, i) = "" Then
                            strYEAR = Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 3, 2)
                            strNEWPROJECTNO = SelectRtn_SEQNO(strYEAR)
                            intRtn = InsertRtn_Sheet_PONO(vntData, intColCnt, i, strNEWPROJECTNO, strCREDAY)
                            strPROJECTNO = strNEWPROJECTNO
                        Else
                            strPROJECTNO = GetElement(vntData, "PROJECTNO", intColCnt, i)
                            intRtn = UpdateRtn_PONO(vntData, intColCnt, i, strCREDAY, strGROUPGBN)
                        End If


                        If i = 1 Then
                            strPROJECTLIST = "'" + strPROJECTNO + "'"
                        Else
                            strPROJECTLIST = strPROJECTLIST + "," + "'" + strPROJECTNO + "'"
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
                mobjcePD_PONO.Dispose()
            End Try
        End With
    End Function

    '============== ProcessRtn PROJECT 코드 수정
    Public Function ProcessRtnSheet(ByVal strInfoXML As String, _
                                    ByVal vntData As Object) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCREDAY As String
        Dim strGROUPGBN
        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_PONO = New cePD_PONO(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        If GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strCREDAY = GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)
                        If GetElement(vntData, "GROUPGBN", intColCnt, i, OPTIONAL_STR) = "그룹" Then
                            strGROUPGBN = "1"
                        Else
                            strGROUPGBN = "2"
                        End If

                        intRtn = UpdateRtn_PONO(vntData, intColCnt, i, strCREDAY, strGROUPGBN)

                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtnSheet")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_PONO.Dispose()
            End Try
        End With
    End Function

    '============== ProcessRtn PROJECT 신규 실제 저장 로직
    Public Function ProcessRtn_Seq(ByVal xmlRoot As XmlElement, _
                                   ByVal strSEQFLAG As String) As Integer
        Dim strPONO
        Dim intRtn
        Dim strCUSTCODE As String
        Dim strYEAR As String
        Dim strCREDAY As String

        With mobjSCGLConfig

            Try
                If XMLGetElement(xmlRoot, "CREDAY") <> "" Then strCREDAY = XMLGetElement(xmlRoot, "CREDAY").Substring(0, 4) & XMLGetElement(xmlRoot, "CREDAY").Substring(5, 2) & XMLGetElement(xmlRoot, "CREDAY").Substring(8, 2)
                ''' 사용할 Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_PONO = New cePD_PONO(mobjSCGLConfig)


                strYEAR = Mid(XMLGetElement(xmlRoot, "CREDAY"), 3, 2)
                If strSEQFLAG = "new" Then
                    strPONO = SelectRtn_SEQNO(strYEAR)
                    intRtn = InsertRtn_PONO(xmlRoot, strPONO, strCREDAY)
                Else
                    'intRtn = UpdateRtn_SEQ(xmlRoot)

                End If

                Return intRtn
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_Seq")
            Finally

                mobjcePD_PONO.Dispose()
            End Try
        End With
    End Function

    '============== ProcessRtn PROJECT CODE 생성
    Public Function SelectRtn_SEQNO(ByVal strYEAR As String) As String
        '여기부터 단순조회
        Dim strSQL, strFormat, strRtn As String
        'SetConfig(strInfoXML) '기본정보 Setting

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = "SELECT 'P'+'" & strYEAR & "'+DBO.LPAD(ISNULL(MAX(CAST(SUBSTRING(PROJECTNO,4,3) AS NUMERIC(4,0))),0)+1,3,'0') FROM PD_PONO WHERE SUBSTRING(PROJECTNO,2,2) = '" & strYEAR & "'"
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
    ' =============== Project Code 삭제여부 판단 조회
    Public Function GetPONODELSELECT(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strCODE As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)



        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            strSQL = "SELECT PROJECTNO FROM PD_JOBNO WHERE PROJECTNO = '" & strCODE & "' AND ENDFLAG <> 'PF01'"
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPONODELSELECT")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    '콤보타입가져오기
    Public Function GetDataType(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strCombo As String) As Object



        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object
        Dim strWhere
        'Combo 구분자 [strCombo] --------------------------------
        'JOBGUBN : 좝구분, JOBBASE : 청구기준, CREGUBN : 제작구분, CREPART : 제작종류, 
        Select Case strCombo
            Case ("JOBGUBN")
                strWhere = "PD_JOBKIND"
            Case ("JOBBASE")
                strWhere = "PD_JOBBASE"
            Case ("CREGUBN")
                strWhere = "PD_CREGUBN"
            Case ("CREPART")
                strWhere = "PD_GRAPHICKIND"
            Case ("ENDFLAG")
                strWhere = "PD_ENDFLAG"
            Case ("PONOGUBN")
                strWhere = "PD_PONOGROUP"
        End Select


        SetConfig(strInfoXML)   '기본정보 설정

        '조회 필드 설정
        strSelFields = "CODE,CODE_NAME"

        'SQL문 생성

        strFormat = "SELECT {0} " & _
                    "FROM SC_CODE " & _
                    "WHERE CLASS_CODE = '" & strWhere & "' " & _
                    "ORDER BY SORT_SEQ "

        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''데이터 조회
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



    '프로젝트 삭제
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strCODE As String) As Integer

        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_PONO = New cePD_PONO(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjcePD_PONO.DeleteDo(strCODE)
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
                mobjcePD_PONO.Dispose()
            End Try
        End With
    End Function
#End Region
    'PROJECTNO,PROJECTNM,CLIENTCODE,CLIENTSUBCODE,SUBSEQ,GROUPGBN,CREDAY,CPDEPTCD,CPEMPNO,MEMO
#Region "GROUP BLOCK : Entity Function Section"
    Private Function InsertRtn_PONO(ByVal xmlRoot As XmlElement, _
                                    ByVal strPONO As String, _
                                    ByVal strCREDAY As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjcePD_PONO.InsertDo( _
                                       strPONO, _
                                       XMLGetElement(xmlRoot, "PROJECTNM"), _
                                       XMLGetElement(xmlRoot, "CLIENTCODE"), _
                                       XMLGetElement(xmlRoot, "SUBSEQ"), _
                                       XMLGetElement(xmlRoot, "GROUPGBN"), _
                                       strCREDAY, _
                                       XMLGetElement(xmlRoot, "CPDEPTCD"), _
                                       XMLGetElement(xmlRoot, "CPEMPNO"), _
                                       XMLGetElement(xmlRoot, "MEMO"), _
                                       XMLGetElement(xmlRoot, "TIMCODE"), _
                                       XMLGetElement(xmlRoot, "ATTR01"), _
                                       XMLGetElement(xmlRoot, "ATTR02"), _
                                       XMLGetElement(xmlRoot, "ATTR03"), _
                                       XMLGetElement(xmlRoot, "ATTR04"), _
                                       XMLGetElement(xmlRoot, "ATTR05"), _
                                       XMLGetElement(xmlRoot, "ATTR06", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "ATTR07", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "ATTR08", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "ATTR09", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "ATTR10", NULL_NUM, True))
        Return intRtn
    End Function


    Private Function InsertRtn_Sheet_PONO(ByVal vntData As Object, _
                                           ByVal intColCnt As Integer, _
                                           ByVal intRow As Integer, _
                                           ByVal strNEWPROJECTNO As String, _
                                           ByVal strCREDAY As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjcePD_PONO.InsertDo( _
                                       strNEWPROJECTNO, _
                                       GetElement(vntData, "PROJECTNM", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "SUBSEQ", intColCnt, intRow), _
                                       GetElement(vntData, "GROUPGBN", intColCnt, intRow), _
                                       strCREDAY, _
                                       GetElement(vntData, "CPDEPTCD", intColCnt, intRow), _
                                       GetElement(vntData, "CPEMPNO", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
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




    Private Function UpdateRtn_PONO(ByVal vntData As Object, _
                                    ByVal intColCnt As Integer, _
                                    ByVal intRow As Integer, _
                                    ByVal strCREDAY As String, _
                                    ByVal strGROUPGBN As String) As Integer
        'PROJECTNO,PROJECTNM,CLIENTCODE,CLIENTSUBCODE,SUBSEQ,GROUPGBN,CREDAY,CPDEPTCD,CPEMPNO,MEMO
        Dim intRtn As Integer
        intRtn = mobjcePD_PONO.UpdateDo( _
                                       GetElement(vntData, "PROJECTNO", intColCnt, intRow), _
                                       GetElement(vntData, "PROJECTNM", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "SUBSEQ", intColCnt, intRow), _
                                       strGROUPGBN, _
                                       strCREDAY, _
                                       GetElement(vntData, "CPDEPTCD", intColCnt, intRow), _
                                       GetElement(vntData, "CPEMPNO", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
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
