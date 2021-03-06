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
Public Class ccPDCOCONTRACT_HADO
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccPDCOCONTRACT_HADO"
    Private mobjcePD_CONTRACT_HDR As ePDCO.cePD_CONTRACT_HDR
    Private mobjcePD_CONTRACT_DTL As ePDCO.cePD_CONTRACT_DTL
    Private mobjcePD_CONTRACT_TEMP As ePDCO.cePD_CONTRACT_TEMP

#End Region

#Region "GROUP BLOCK : Function Section"
    '=============== 계약서 미등록 내역 조회 
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strFROM As String, _
                              ByVal strTO As String, _
                              ByVal strOUTSCODE As String, _
                              ByVal strOUTSNAME As String, _
                              ByVal strJOBNAME As String) As Object

        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim Con1, Con2, Con3, Con4 As String


        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML) '기본정보 Setting

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""

                If strFROM <> "" And strTO <> "" Then
                    Con1 = String.Format(" AND (CONTRACTDAY BETWEEN '{0}' AND  '{1}')", Replace(strFROM, "-", ""), Replace(strTO, "-", ""))
                End If
                If strFROM <> "" And strTO = "" Then
                    Con1 = String.Format(" AND (CONTRACTDAY >= '{0}')", Replace(strFROM, "-", ""))
                End If
                If strFROM = "" And strTO <> "" Then
                    Con1 = String.Format(" AND (CONTRACTDAY <= '{0}')", Replace(strTO, "-", ""))
                End If

                If strOUTSCODE <> "" Then Con2 = String.Format(" AND OUTSCODE = '{0}'", strOUTSCODE)
                If strOUTSNAME <> "" Then Con3 = String.Format(" AND DBO.SC_GET_HIGHCUSTNAME_FUN(OUTSCODE) LIKE '%{0}%'", strOUTSNAME)

                strJOBNAME = Replace(strJOBNAME, "'", "''")
                If strJOBNAME <> "" Then Con4 = String.Format(" AND DBO.PD_JOBNAME_FUN(JOBNO) like '%{0}%'", strJOBNAME)


                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strFormat = " SELECT "
                strFormat = strFormat & " 0 CHK, "
                strFormat = strFormat & " SEQ, "
                strFormat = strFormat & " OUTSCODE, "
                strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(OUTSCODE) OUTSNAME, "
                strFormat = strFormat & " CONTRACTDAY,"
                strFormat = strFormat & " DBO.PD_JOBNAME_FUN(JOBNO) JOBNAME, "
                strFormat = strFormat & " JOBNO, "
                strFormat = strFormat & " AMT,"
                strFormat = strFormat & " MEMO"
                strFormat = strFormat & " FROM PD_CONTRACT_DTL"
                strFormat = strFormat & " where 1=1 {0} AND ISNULL(CONTRACTNO,'') = '' "
                strFormat = strFormat & " ORDER BY OUTSCODE"

                strSQL = String.Format(strFormat, strWhere)

                ' DB 접속
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


    '=============== 계약서 등록 내역 조회
    Public Function SelectRtn_EXIST(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strFROM As String, _
                                    ByVal strTO As String, _
                                    ByVal strOUTSCODE As String, _
                                    ByVal strOUTSNAME As String, _
                                    ByVal strJOBNAME As String, _
                                    ByVal strCONFIRM As String, _
                                    ByVal strCONTRACTNO As String, _
                                    ByVal strCONTRACTNAME As String) As Object


        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7 As String

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체

            Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = ""

            Try
                If strFROM <> "" And strTO <> "" Then
                    Con1 = String.Format(" AND (CONTRACTDAY BETWEEN '{0}' AND  '{1}')", Replace(strFROM, "-", ""), Replace(strTO, "-", ""))
                End If
                If strFROM <> "" And strTO = "" Then
                    Con1 = String.Format(" AND (CONTRACTDAY >= '{0}')", Replace(strFROM, "-", ""))
                End If
                If strFROM = "" And strTO <> "" Then
                    Con1 = String.Format(" AND (CONTRACTDAY <= '{0}')", Replace(strTO, "-", ""))
                End If

                If strOUTSCODE <> "" Then Con2 = String.Format(" AND OUTSCODE = '{0}'", strOUTSCODE)
                If strOUTSNAME <> "" Then Con3 = String.Format(" AND DBO.SC_GET_HIGHCUSTNAME_FUN(OUTSCODE) LIKE '%{0}%'", strOUTSNAME)

                If strJOBNAME <> "" Then Con4 = String.Format(" AND CONTRACTNO IN(SELECT CONTRACTNO FROM PD_CONTRACT_DTL WHERE DBO.PD_JOBNAME_FUN(JOBNO) like '%{0}%')", strJOBNAME)
                If strCONFIRM <> "" Then Con5 = String.Format(" AND CONFIRMFLAG = '{0}'", strCONFIRM)
                If strCONTRACTNO <> "" Then Con6 = String.Format(" AND CONTRACTNO LIKE '%{0}%'", strCONTRACTNO)
                If strCONTRACTNAME <> "" Then Con7 = String.Format(" AND CONTRACTNAME LIKE '%{0}%'", strCONTRACTNAME)


                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7)

                strFormat = " SELECT "
                strFormat = strFormat & " 0 CHK ,  "
                strFormat = strFormat & " CONTRACTNO, "
                strFormat = strFormat & " CONTRACTNAME, "
                strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(OUTSCODE) OUTSNAME, "
                strFormat = strFormat & " CONTRACTDAY, "
                strFormat = strFormat & " LOCALAREA, "
                strFormat = strFormat & " STDATE, "
                strFormat = strFormat & " EDDATE, "
                strFormat = strFormat & " AMT, "
                strFormat = strFormat & " DELIVERYDAY, "
                strFormat = strFormat & " TESTDAY, "
                strFormat = strFormat & " PAYMENTGBN, "
                strFormat = strFormat & " TESTMENT, "
                strFormat = strFormat & " COMENT, "
                strFormat = strFormat & " CONFIRMFLAG, "
                strFormat = strFormat & " PRERATE, "
                strFormat = strFormat & " PREAMT, "
                strFormat = strFormat & " ENDRATE, "
                strFormat = strFormat & " ENDAMT, "
                strFormat = strFormat & " THISRATE, "
                strFormat = strFormat & " THISAMT, "
                strFormat = strFormat & " BALANCERATE, "
                strFormat = strFormat & " BALANCEAMT, "
                strFormat = strFormat & " DELIVERYGUARANTY, "
                strFormat = strFormat & " FAULTGUARANTY, "
                strFormat = strFormat & " MANAGER, "
                strFormat = strFormat & " TESTENDDAY, "
                strFormat = strFormat & " TESTAMT, "
                strFormat = strFormat & " LOSTDAY, "
                strFormat = strFormat & " CONFLAG, "
                strFormat = strFormat & " DIVFLAG "
                strFormat = strFormat & " FROM PD_CONTRACT_HDR "
                strFormat = strFormat & " WHERE 1=1 {0}"
                strFormat = strFormat & " ORDER BY OUTSCODE,CONTRACTNO"

                strSQL = String.Format(strFormat, strWhere)

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
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_EXIST")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function



    ' ============== 계약서 저장
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object, _
                               ByVal strENDFLAG As String) As Object

        Dim intRtn As Integer '결과값 변수
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수 
        Dim strSEQ
        Dim strCONTRACTDAY
        Dim strCONTRACTNO
        Dim strCONFIRMFLAG

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_CONTRACT_DTL = New cePD_CONTRACT_DTL(mobjSCGLConfig)
                    mobjcePD_CONTRACT_HDR = New cePD_CONTRACT_HDR(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        If strENDFLAG = "F" Then
                            strCONTRACTDAY = ""
                            If GetElement(vntData, "CONTRACTDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strCONTRACTDAY = GetElement(vntData, "CONTRACTDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "CONTRACTDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "CONTRACTDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)

                            If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                                intRtn = InsertRtn_DTL(vntData, intColCnt, i, strCONTRACTDAY)
                            Else
                                strSEQ = GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True)
                                intRtn = UpdateRtn_DTL(vntData, intColCnt, i, strCONTRACTDAY)
                            End If
                        Else
                            strCONTRACTNO = GetElement(vntData, "CONTRACTNO", intColCnt, i, OPTIONAL_STR)

                            If GetElement(vntData, "CONFIRMFLAG", intColCnt, i, OPTIONAL_STR) = 1 Then
                                strCONFIRMFLAG = "1"
                            Else
                                strCONFIRMFLAG = "0"

                            End If

                            intRtn = UpdateRtn_CONFIRM(vntData, intColCnt, i, strCONTRACTNO, strCONFIRMFLAG)

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
                mobjcePD_CONTRACT_DTL.Dispose()
                mobjcePD_CONTRACT_HDR.Dispose()
            End Try
        End With
    End Function

    Public Function ProcessRtn_HDR(ByVal strInfoXML As String, _
                                   ByVal strMasterXML As String, _
                                   ByVal vntData As Object, _
                                   ByRef strCONTRACTNO As String, _
                                   ByRef strOUTSCODE As String) As Integer

        Dim intRtn As Integer
        Dim intCnt As Integer
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수 
        Dim strSQL As String
        Dim strCONFIRMFLAG
        Dim strSTDATE, strEDDATE, strCONTRACTDAY, strDELIVERYDAY, strTESTDAY, strTESTENDDAY, strLOSTDAY
        Dim strCOMENT
        Dim strCONFLAG, strDIVFLAG

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                Dim xmlRoot As XmlElement
                xmlRoot = XMLGetRoot(strMasterXML) 'XML 데이터

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                ''' 사용할 Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_CONTRACT_HDR = New cePD_CONTRACT_HDR(mobjSCGLConfig)
                mobjcePD_CONTRACT_DTL = New cePD_CONTRACT_DTL(mobjSCGLConfig)

                strDELIVERYDAY = "" : strSTDATE = "" : strEDDATE = "" : strCONTRACTDAY = "" : strTESTDAY = ""
                strTESTENDDAY = "" : strLOSTDAY = "" : strCOMENT = "" : strCONFLAG = "" : strDIVFLAG = ""

                If XMLGetElement(xmlRoot, "DELIVERYDAY") <> "" Then strDELIVERYDAY = Replace(XMLGetElement(xmlRoot, "DELIVERYDAY"), "-", "")
                If XMLGetElement(xmlRoot, "STDATE") <> "" Then strSTDATE = Replace(XMLGetElement(xmlRoot, "STDATE"), "-", "")
                If XMLGetElement(xmlRoot, "EDDATE") <> "" Then strEDDATE = Replace(XMLGetElement(xmlRoot, "EDDATE"), "-", "")
                If XMLGetElement(xmlRoot, "CONTRACTDAY") <> "" Then strCONTRACTDAY = Replace(XMLGetElement(xmlRoot, "CONTRACTDAY"), "-", "")
                If XMLGetElement(xmlRoot, "TESTDAY") <> "" Then strTESTDAY = Replace(XMLGetElement(xmlRoot, "TESTDAY"), "-", "")
                If XMLGetElement(xmlRoot, "TESTENDDAY") <> "" Then strTESTENDDAY = Replace(XMLGetElement(xmlRoot, "TESTENDDAY"), "-", "")
                If XMLGetElement(xmlRoot, "LOSTDAY") <> "" Then strLOSTDAY = Replace(XMLGetElement(xmlRoot, "LOSTDAY"), "-", "")

                strCOMENT = XMLGetElement(xmlRoot, "COMENT")

                If XMLGetElement(xmlRoot, "CONFIRMFLAG") <> "" Then
                    If XMLGetElement(xmlRoot, "CONFIRMFLAG") = "1" Or XMLGetElement(xmlRoot, "CONFIRMFLAG") = "-1" Then
                        strCONFIRMFLAG = "1"
                    Else
                        strCONFIRMFLAG = "0"
                    End If
                Else
                    strCONFIRMFLAG = "0"
                End If

                If XMLGetElement(xmlRoot, "CONFLAG") <> "" Then
                    If XMLGetElement(xmlRoot, "CONFLAG") = "1" Or XMLGetElement(xmlRoot, "CONFLAG") = "-1" Then
                        strCONFLAG = "1"
                    Else
                        strCONFLAG = "0"
                    End If
                Else
                    strCONFLAG = "0"
                End If

                If XMLGetElement(xmlRoot, "DIVFLAG") <> "" Then
                    If XMLGetElement(xmlRoot, "DIVFLAG") = "1" Or XMLGetElement(xmlRoot, "DIVFLAG") = "-1" Then
                        strDIVFLAG = "1"
                    Else
                        strDIVFLAG = "0"
                    End If
                Else
                    strDIVFLAG = "0"
                End If

                strCONTRACTNO = SelectRtn_CONTRACTSEQNO(strCONTRACTDAY)


                intRtn = InsertRtn_HDR(xmlRoot, strCONTRACTNO, strOUTSCODE, strDELIVERYDAY, _
                                       strSTDATE, strEDDATE, strCONTRACTDAY, strTESTDAY, strTESTENDDAY, _
                                       strLOSTDAY, strCOMENT, strCONFLAG, strDIVFLAG)


                '여기부터는 PD_CONTRACT_DTL 에 CONTRACTNO 업데이트 시작
                intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                Dim strSEQ

                For i = 1 To intRows
                    If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then
                        strSEQ = GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True)

                        intRtn = UpdateRtn_DTL_CONTRACTNO(vntData, intColCnt, i, strSEQ, strCONTRACTNO)

                    End If
                Next
                'PD_CONTRACT_DTL 에 CONTRACTNO 업데이트 끝

                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_HDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjcePD_CONTRACT_DTL.Dispose()
                mobjcePD_CONTRACT_HDR.Dispose()
            End Try
        End With
    End Function

    '============== ProcessRtn 계약서 CONTRACTCODE 생성
    Public Function SelectRtn_CONTRACTSEQNO(ByVal strCONTRACTDAY As String) As String
        '여기부터 단순조회
        Dim strSQL, strFormat, strRtn As String
        'SetConfig(strInfoXML) '기본정보 Setting

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = "SELECT '판관' + '" & strCONTRACTDAY & "'+'-'+dbo.lpad(ISNULL(max(cast(right(CONTRACTNO,3) as numeric)),0)+1,3,0)  "
                strSQL = strSQL & " FROM PD_CONTRACT_HDR "
                strSQL = strSQL & " WHERE SUBSTRING(CONTRACTNO,1,10) = '판관' + '" & strCONTRACTDAY & "'"

                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CONTRACTSEQNO")
            Finally
            End Try
        End With
        '여기까지 단순조회
    End Function

    '================계약서 삭제
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strSEQ As String) As Integer

        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_CONTRACT_DTL = New cePD_CONTRACT_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjcePD_CONTRACT_DTL.DeleteDo(strSEQ)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                '사용한 Entity(개체Dispose)
                mobjcePD_CONTRACT_DTL.Dispose()
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    Public Function DeleteRtn_HDR(ByVal strInfoXML As String, _
                                  ByVal vntData As Object) As Object

        Dim intRtn As Integer '결과값 변수
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수 
        Dim strCONTRACTNO

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_CONTRACT_HDR = New cePD_CONTRACT_HDR(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        strCONTRACTNO = ""
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then
                            strCONTRACTNO = GetElement(vntData, "CONTRACTNO", intColCnt, i, OPTIONAL_STR)

                            intRtn = mobjcePD_CONTRACT_HDR.DeleteRtn_Confirm_HDR(strCONTRACTNO)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".DeleteRtn_HDR")
            Finally
                mobjcePD_CONTRACT_HDR.Dispose()
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' ============== ProcessRtn_TEMP (Master & Detail) Sample Code 
    Public Function ProcessRtn_TEMP(ByVal strInfoXML As String, _
                                    ByRef strCONTACTNO As String, _
                                    ByRef dblNUM As Double, _
                                    ByRef strUSERID As String) As Integer

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim intCnt

        SetConfig(strInfoXML) '기본정보 Setting

        With mobjSCGLConfig '기본정보를 가지고 있는 Config 개체
            Try
                'DB접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master 데이터 처리
                mobjcePD_CONTRACT_TEMP = New cePD_CONTRACT_TEMP(mobjSCGLConfig)

                strUSERID = .WRKUSR 'USERID 를 받아야 출력시에 USERID에 따른 구분을 둘수있다.
                intRtn = InsertRtn_CONTRACT_TEMP(strCONTACTNO, dblNUM, strUSERID)

                '트랜잭션Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_TEMP")
            Finally
                'Resource해제
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjcePD_CONTRACT_TEMP.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn_temp(ByVal strInfoXML As String) As Integer   '데이터 DELETE

        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_CONTRACT_TEMP = New cePD_CONTRACT_TEMP(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjcePD_CONTRACT_TEMP.DeleteDo(.WRKUSR)
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
                mobjcePD_CONTRACT_TEMP.Dispose()
            End Try
        End With
    End Function
#End Region


#Region "GROUP BLOCK : Entity Function Section"
    Private Function InsertRtn_HDR(ByVal xmlRoot As XmlElement, _
                                   ByVal strCONTRACTNO As String, _
                                   ByVal strOUTSCODE As String, _
                                   ByVal strDELIVERYDAY As String, _
                                   ByVal strSTDATE As String, _
                                   ByVal strEDDATE As String, _
                                   ByVal strCONTRACTDAY As String, _
                                   ByVal strTESTDAY As String, _
                                   ByVal strTESTENDDAY As String, _
                                   ByVal strLOSTDAY As String, _
                                   ByVal strCOMENT As String, _
                                   ByVal strCONFLAG As String, _
                                   ByVal strDIVFLAG As String) As Integer


        Dim intRtn As Integer
        intRtn = mobjcePD_CONTRACT_HDR.InsertDo( _
                                       strCONTRACTNO, _
                                       strOUTSCODE, _
                                       XMLGetElement(xmlRoot, "CONTRACTNAME"), _
                                       XMLGetElement(xmlRoot, "LOCALAREA"), _
                                       strDELIVERYDAY, _
                                       XMLGetElement(xmlRoot, "AMT", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "PRERATE", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "PREAMT", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "ENDRATE", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "ENDAMT", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "THISRATE", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "THISAMT", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "BALANCERATE", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "BALANCEAMT", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "DELIVERYGUARANTY", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "FAULTGUARANTY", NULL_NUM, True), _
                                       XMLGetElement(xmlRoot, "PAYMENTGBN"), _
                                       strSTDATE, _
                                       strEDDATE, _
                                       strCONTRACTDAY, _
                                       XMLGetElement(xmlRoot, "MANAGER"), _
                                       strTESTDAY, _
                                       strTESTENDDAY, _
                                       XMLGetElement(xmlRoot, "TESTMENT"), _
                                       XMLGetElement(xmlRoot, "TESTAMT", NULL_NUM, True), _
                                       strLOSTDAY, _
                                       XMLGetElement(xmlRoot, "CONFIRMFLAG"), _
                                       strCONFLAG, _
                                       strDIVFLAG, _
                                       strCOMENT, _
                                       XMLGetElement(xmlRoot, "AMTFLAG"), _
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

    Private Function InsertRtn_DTL(ByVal vntData As Object, _
                                   ByVal intColCnt As Integer, _
                                   ByVal intRow As Integer, _
                                   ByVal strCONTRACTDAY As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjcePD_CONTRACT_DTL.InsertDo( _
                                       GetElement(vntData, "OUTSCODE", intColCnt, intRow), _
                                       GetElement(vntData, "JOBNO", intColCnt, intRow), _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       strCONTRACTDAY, _
                                       GetElement(vntData, "CONTRACTNO", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "CONFIRM_USER", intColCnt, intRow), _
                                       GetElement(vntData, "CONFIRM_DATE", intColCnt, intRow), _
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

    Private Function UpdateRtn_DTL(ByVal vntData As Object, _
                                   ByVal intColCnt As Integer, _
                                   ByVal intRow As Integer, _
                                   ByVal strCONTRACTDAY As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjcePD_CONTRACT_DTL.UpdateDo( _
                                       GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "OUTSCODE", intColCnt, intRow), _
                                       GetElement(vntData, "JOBNO", intColCnt, intRow), _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       strCONTRACTDAY, _
                                       GetElement(vntData, "CONTRACTNO", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "CONFIRM_USER", intColCnt, intRow), _
                                       GetElement(vntData, "CONFIRM_DATE", intColCnt, intRow), _
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

    Private Function UpdateRtn_DTL_CONTRACTNO(ByVal vntData As Object, _
                                              ByVal intColCnt As Integer, _
                                              ByVal intRow As Integer, _
                                              ByVal strSEQ As Integer, _
                                              ByVal strCONTRACTNO As String) As Integer


        Dim intRtn As Integer
        intRtn = mobjcePD_CONTRACT_DTL.UpdateDo_CONTRACTNO(strSEQ, _
                                                           strCONTRACTNO)

        Return intRtn
    End Function

    Private Function InsertRtn_CONTRACT_TEMP(ByVal strCONTRACTNO As String, _
                                             ByRef dblNUM As String, _
                                             ByRef strUSERID As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjcePD_CONTRACT_TEMP.InsertDo( _
                                       strCONTRACTNO, _
                                       dblNUM, _
                                       strUSERID)
        Return intRtn

    End Function


    Private Function UpdateRtn_CONFIRM(ByVal vntData As Object, _
                                       ByVal intColCnt As Integer, _
                                       ByVal intRow As Integer, _
                                       ByVal strCONTRACTNO As String, _
                                       ByVal strCONFIRMFLAG As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjcePD_CONTRACT_HDR.UpdateDo_CONFIRM( _
                                                        strCONTRACTNO, _
                                                        strCONFIRMFLAG)

        Return intRtn
    End Function

#End Region
End Class

