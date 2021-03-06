'****************************************************************************************
'Generated By  : Kim Tae Ho 
'시스템구분    : RMS/PD/Server Control Class
'실행   환경   : COM+ Service Server Package
'프로그램명    : ccPDCMPREESTLIST.vb
'기         능 : - 가견적관리
'특이  사항    : - CE 단 Query 복사 기능
'                -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008.11.12 Kim Tae Ho
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
Public Class ccPDCOEXEENDLIST
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccPDCOEXEENDLIST"                  '자신의 클래스명
    Private mobjcePD_EXE_HDR As ePDCO.cePD_EXE_HDR            '외주정산 기본내역
    Private mobjcePD_EXE_DTL As ePDCO.cePD_EXE_DTL            '외주정산 상세내역
    Private mobjcePD_ACC_MST As ePDCO.cePD_ACC_MST            '회계진행비 상세내역 투입
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password"
#End Region

#Region "GROUP BLOCK : Function Section"
    ' =============== 정산마감시 해당 JOB 조회
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strFROM As String, _
                              ByVal strTO As String, _
                              ByVal strPROJECTNO As String, _
                              ByVal strPROJECTNM As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCLIENTNAME As String, _
                              ByVal strGUBN As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)
        Dim Con1, Con2, Con3, Con4, Con5, Con6
        'Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,cmbSEARCHENDFLAG.value
        strCols = " A.PROJECTNO, "
        strCols = strCols & " A.JOBNO,"
        strCols = strCols & " A.JOBNAME,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,"
        strCols = strCols & " A.CLIENTSUBCODE,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(A.CLIENTSUBCODE) CLIENTSUBNAME,"
        strCols = strCols & " A.SUBSEQ,"
        strCols = strCols & " DBO.PD_JOBCUST_NAME_FUN(A.SUBSEQ) SUBSEQNAME,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.ENDFLAG) ENDFLAG,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.JOBGUBN) JOBGUBN,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.CREPART) CREPART,"
        strCols = strCols & " DBO.MD_GETPUBNAME_FUN(A.CREGUBN) CREGUBN,"
        strCols = strCols & " REQDAY,DBO.PD_COMMITION_FUN(A.CLIENTCODE) COMMITION,A.CLIENTCODE,A.PREESTNO,"
        strCols = strCols & " CASE B.DIVAMT-B.AMT WHEN 0 THEN '청구완료' ELSE '청구미완료' END DIVFLAG,"
        strCols = strCols & " C.ENDDAY,B.DIVAMT,B.AMT, A.DEMANDYEARMON"

        If strFROM <> "" And strTO <> "" Then
            Con1 = String.Format(" AND (A.CREDAY BETWEEN '{0}' AND  '{1}')", strFROM, strTO)
        End If
        If strFROM <> "" And strTO = "" Then
            Con1 = String.Format(" AND (A.CREDAY > '{0}')", strFROM)
        End If
        If strFROM = "" And strTO <> "" Then
            Con1 = String.Format(" AND (A.CREDAY < '{0}')", strTO)
        End If


        If strPROJECTNM <> "" Then Con2 = String.Format(" AND (A.PROJECTNM like '%{0}%')", strPROJECTNM)
        If strPROJECTNO <> "" Then Con3 = String.Format(" AND (A.PROJECTNO = '{0}')", strPROJECTNO)
        If strCLIENTCODE <> "" Then Con4 = String.Format(" AND (A.CLIENTCODE = '{0}')", strCLIENTCODE)
        If strCLIENTNAME <> "" Then Con5 = String.Format(" AND (LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) like '%{0}%')", strCLIENTNAME)
        If strGUBN <> "" Then Con6 = String.Format(" AND (CASE ISNULL(C.ENDDAY,'') WHEN '' THEN 'F' ELSE 'T' END = '{0}')", strGUBN)

        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6)

        strFormat = "SELECT {0} FROM PD_EXE_HDR C,V_JOBNO A LEFT JOIN (SELECT X.JOBNO,ISNULL(X.DIVAMT,0) DIVAMT,ISNULL(Y.AMT,0) AMT FROM (SELECT JOBNO,SUM(DIVAMT) DIVAMT FROM PD_DIVAMT GROUP BY JOBNO) X LEFT JOIN (SELECT JOBNO,SUM(AMT) AMT FROM PD_TAX_MST GROUP BY JOBNO) Y ON X.JOBNO = Y.JOBNO) B ON A.JOBNO = B.JOBNO WHERE A.JOBNO = C.JOBNO {1} ORDER BY A.JOBNO DESC"

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
    ' =============== 정산마감 JOB별 정산내역
    Public Function SelectRtn_DTL(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strCODE As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)
        Dim Con1

        strCols = " DBO.PD_ESTJOBNO_FUN(A.PREESTNO) JOBNO,"
        strCols = strCols & " A.PREESTNO,"
        strCols = strCols & " CASE ISNULL(B.SORTSEQ,0) WHEN 0 THEN rank() OVER (ORDER BY A.ITEMCODESEQ) ELSE B.SORTSEQ END AS SORTSEQ,"
        strCols = strCols & " A.ITEMCODESEQ,"
        strCols = strCols & " A.ITEMCODE,"
        strCols = strCols & " DBO.PD_ITEMDIVNAME_FUN(A.ITEMCODE) ITEMCLASS,"
        strCols = strCols & " DBO.PD_ITEMCODENAME_FUN(A.ITEMCODE) ITEMNAME,"
        strCols = strCols & " CASE ISNULL(B.ADDFLAG,'') WHEN '' THEN A.QTY ELSE Null END AS QTY,"
        strCols = strCols & " CASE ISNULL(B.ADDFLAG,'') WHEN '' THEN A.PRICE ELSE Null END AS PRICE,"
        strCols = strCols & " CASE ISNULL(B.ADDFLAG,'') WHEN '' THEN A.AMT ELSE Null END AS AMT,"
        strCols = strCols & " B.OUTSCODE,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(B.OUTSCODE) OUTSNAME,"
        strCols = strCols & " B.ADJAMT,"
        strCols = strCols & " B.STD,"
        strCols = strCols & " B.VOCHNO,"
        strCols = strCols & " B.ADJDAY,B.ADDFLAG,B.SEQ,B.PURCHASENO"


        If strCODE <> "" Then Con1 = String.Format(" AND (DBO.PD_ESTJOBNO_FUN(A.PREESTNO) = '{0}')", strCODE)
        strWhere = BuildFields(" ", Con1)
        strFormat = "SELECT {0} FROM PD_PREEST_HDR C,PD_PREEST_DTL A LEFT JOIN PD_EXE_DTL B ON A.PREESTNO = B.PREESTNO AND A.ITEMCODESEQ = B.ITEMCODESEQ AND A.ITEMCODE = B.ITEMCODE "
        strFormat = strFormat & "WHERE  C.PREESTNO = A.PREESTNO AND C.JOBNO = '" & strCODE & "' AND ISNULL(C.CONFIRMFLAG,'') <> '' {1} ORDER BY CASE ISNULL(B.SORTSEQ,0) WHEN 0 THEN rank() OVER (ORDER BY A.ITEMCODESEQ) ELSE B.SORTSEQ END"

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData

            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_DTL")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '정산마감 최초저장처리
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer '결과값 변수
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보를 가지고 있는 Config 개체
            Try
                'XML Element 변수 선언 (strMasterXML을 변환)
                Dim xmlRoot As XmlElement


                'DB접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master 데이터 처리


                'Detail 데이터 처리
                If IsArray(vntData) Then
                    intRtn = ProcessRtn_DTL(vntData)
                End If
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
    '============== 정산마감처리 
    Public Function ProcessRtn_DTL(ByVal vntData As Object) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim lngSEQ As Double
        Dim strPURCHASENO As String
        Dim strPRECODE As String
        Dim strENDDAY As String
        Dim strGROUPGBN
        Dim strJOBNO As String
        Dim strSEQ As Double
        Dim strSQLEND As String
        Dim strSQL As String
        Dim strENDSQL As String
        Dim intRtn2 As Integer
        With mobjSCGLConfig
            Try

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_EXE_HDR = New cePD_EXE_HDR(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)


                    For i = 1 To intRows
                        '인서트
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then
                            If GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strENDDAY = GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "ENDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)

                            strJOBNO = GetElement(vntData, "JOBNO", intColCnt, i, OPTIONAL_STR)

                            '정산 헤더에 마감일자를 업데이트 한다.
                            strSQL = "UPDATE PD_EXE_HDR SET ENDDAY = '" & strENDDAY & "' WHERE JOBNO = '" & strJOBNO & "'"
                            '[변경]
                            'intRtn = mobjcePD_EXE_HDR.UpdateRtn_Endday(strSQL)

                            '결산 자료를 생성 한다. << - 정산마감 자료 생성은 일괄 처리 로 변경 (PD_CLOSING_MST)
                            'intRtn2 = mobjcePD_EXE_HDR.InsertClosing(strJOBNO, strENDDAY)

                            '마감처리시 JOB 번호 의 STATUS(상태) 값을 PF04[결산상태] 로 변경
                            strSQLEND = "UPDATE PD_JOBNO SET ENDFLAG = 'PF04',SETYEARMON = '" & strENDDAY & "' WHERE JOBNO = '" & strJOBNO & "'"
                            intRtn = mobjcePD_EXE_HDR.ENDFLAG_Update(strSQLEND)
                        End If
                    Next
                End If

                Return intRows
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_DTL")
            Finally

                mobjcePD_EXE_HDR.Dispose()
            End Try
        End With
    End Function
    '정산마감취소 최초저장처리
    Public Function ProcessRtn_Cancel(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer '결과값 변수
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보를 가지고 있는 Config 개체
            Try
                'XML Element 변수 선언 (strMasterXML을 변환)
                Dim xmlRoot As XmlElement


                'DB접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master 데이터 처리

                'Detail 데이터 처리
                If IsArray(vntData) Then
                    intRtn = ProcessRtn_Cancel_DTL(vntData)
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_Cancel")
            Finally
                'Resource해제
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '============== 정산마감취소 처리  
    Public Function ProcessRtn_Cancel_DTL(ByVal vntData As Object) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim lngSEQ As Double
        Dim strPURCHASENO As String
        Dim strPRECODE As String
        Dim strADJDAY As String
        Dim strGROUPGBN
        Dim strJOBNO As String
        Dim strSEQ As Double
        Dim strSQL As String
        Dim strSQLEND As String
        With mobjSCGLConfig
            Try

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_EXE_HDR = New cePD_EXE_HDR(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        '인서트
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then

                            strJOBNO = GetElement(vntData, "JOBNO", intColCnt, i, OPTIONAL_STR)
                            strSQL = "UPDATE PD_EXE_HDR SET ENDDAY = '' WHERE JOBNO = '" & strJOBNO & "'"
                            '[변경]
                            'intRtn = mobjcePD_EXE_HDR.UpdateRtn_Endday(strSQL)
                            '마감취소처리 시 JOB 번호 의 STATUS(상태) 값을 PF03 [청구상태로]로 변경
                            strSQLEND = "UPDATE PD_JOBNO SET ENDFLAG = 'PF03',SETYEARMON = '' WHERE JOBNO = '" & strJOBNO & "'"
                            intRtn = mobjcePD_EXE_HDR.ENDFLAG_Update(strSQLEND)
                        End If
                    Next
                End If

                Return intRows
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_Cancel_DTL")
            Finally

                mobjcePD_EXE_HDR.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : Entity Function Section"




#End Region
End Class

