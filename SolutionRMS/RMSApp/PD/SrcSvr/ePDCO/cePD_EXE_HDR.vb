'****************************************************************************************
'시스템구분 : RMS/PD/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : cePD_JOBNO.vb (PD_JOBNO Entity 처리 Class)
'기      능 : PD_JOBNO Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-11-07 
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class cePD_EXE_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_EXE_HDR"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    'JOBNO,DEMANDDAY,ENDDAY,PREESTNO,NONCOMMITION,COMMITION,SUSURATE,SUSUAMT,DEMANDAMT,PAYMENT,RATE,INCOM
    Public Function InsertDo(ByVal strJOBNO As String, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal dblNONCOMMITION As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMITION As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblDEMANDAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPAYMENT As Double = OPTIONAL_NUM, _
            Optional ByVal dblRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblINCOM As Double = OPTIONAL_NUM, _
            Optional ByVal dblESTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "ENDDAY", strENDDAY, strFields, strValues)
            BuildNameValues(",", "PREESTNO", strPREESTNO, strFields, strValues)
            BuildNameValues(",", "NONCOMMITION", dblNONCOMMITION, strFields, strValues)
            BuildNameValues(",", "COMMITION", dblCOMMITION, strFields, strValues)
            BuildNameValues(",", "SUSURATE", dblSUSURATE, strFields, strValues)
            BuildNameValues(",", "SUSUAMT", dblSUSUAMT, strFields, strValues)
            BuildNameValues(",", "DEMANDAMT", dblDEMANDAMT, strFields, strValues)
            BuildNameValues(",", "PAYMENT", dblPAYMENT, strFields, strValues)
            BuildNameValues(",", "RATE", dblRATE, strFields, strValues)
            BuildNameValues(",", "INCOM", dblINCOM, strFields, strValues)
            BuildNameValues(",", "ESTAMT", dblESTAMT, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", dblATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", dblATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", dblATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", dblATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", dblATTR10, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    'JOBNO,DEMANDDAY,ENDDAY,PREESTNO,NONCOMMITION,COMMITION,SUSURATE,SUSUAMT,DEMANDAMT,PAYMENT,RATE,INCOM
    Public Function UpdateDo(ByVal strJOBNO As String, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPREESTNO As String = OPTIONAL_STR, _
            Optional ByVal dblNONCOMMITION As Double = OPTIONAL_NUM, _
            Optional ByVal dblCOMMITION As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSURATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblSUSUAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblDEMANDAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPAYMENT As Double = OPTIONAL_NUM, _
            Optional ByVal dblRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblINCOM As Double = OPTIONAL_NUM, _
            Optional ByVal dblESTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM)

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("JOBNO", strJOBNO), _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("ENDDAY", strENDDAY), _
                        GetFieldNameValue("PREESTNO", strPREESTNO), _
                        GetFieldNameValue("NONCOMMITION", dblNONCOMMITION), _
                        GetFieldNameValue("COMMITION", dblCOMMITION), _
                        GetFieldNameValue("SUSURATE", dblSUSURATE), _
                        GetFieldNameValue("SUSUAMT", dblSUSUAMT), _
                        GetFieldNameValue("DEMANDAMT", dblDEMANDAMT), _
                        GetFieldNameValue("PAYMENT", dblPAYMENT), _
                        GetFieldNameValue("RATE", dblRATE), _
                        GetFieldNameValue("INCOM", dblINCOM), _
                        GetFieldNameValue("ESTAMT", dblESTAMT), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("ATTR04", strATTR04), _
                        GetFieldNameValue("ATTR05", strATTR05), _
                        GetFieldNameValue("ATTR06", dblATTR06), _
                        GetFieldNameValue("ATTR07", dblATTR07), _
                        GetFieldNameValue("ATTR08", dblATTR08), _
                        GetFieldNameValue("ATTR09", dblATTR09), _
                        GetFieldNameValue("ATTR10", dblATTR10), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("JOBNO", strJOBNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function


    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(ByVal strCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_PREEST_DTL WHERE PREESTNO = '" & strCODE & "'; DELETE FROM PD_PREEST_HDR WHERE PREESTNO = '" & strCODE & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Select 처리
    '*****************************************************************
    Public Function SelectDo(ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                            Optional ByVal strSEQNO As String = OPTIONAL_STR, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("SEQNO", strSEQNO))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리
    '*****************************************************************
    Public Function ENDFLAG_Update(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ENDFLAG_Update")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 제작번호 의 ENDFLAG 를 변경 한다 의뢰 >> 진행
    '*****************************************************************
    Public Function UpdateENDFLAG(ByVal strJOBNO As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_JOBNO SET ENDFLAG = 'PF02' WHERE JOBNO = '" & strJOBNO & "' "
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateENDFLAG")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 정산완료 마감 및 마감취소처리 
    '*****************************************************************
    Public Function Update_ENDDAY(ByVal strJOBNO As String, ByVal strNOW As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_EXE_HDR SET ENDDAY = '" & strNOW & "' WHERE JOBNO = '" & strJOBNO & "' "
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_Endday")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 정산완료 마감 및 마감취소처리 
    '*****************************************************************
    Public Function Delete_ENDDAY(ByVal strJOBNO As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE PD_EXE_HDR SET ENDDAY = '' WHERE JOBNO = '" & strJOBNO & "' "
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_Endday")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 외주정산 최초저장시 청구금액을 외주정산 헤더에 투입
    '*****************************************************************
    Public Function UpdateRtn_DEMANDAMT(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_Demandamt")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 진행비가 입력될때마다 금액을 외주정산 헤더에 업데이트
    '*****************************************************************
    Public Function UpdateRtn_ACCAMT(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_Demandamt")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 마감처리시 결산데이터 생성(미지급금) 아래의 로직으로 대처하였으므로 더이상 사용을 금함!
    '*****************************************************************
    Public Function InsertClosing(ByVal strJOBNO As String, _
                                  ByVal strENDDAY As String) As Integer
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = " INSERT INTO PD_CLOSING_MST "
            strSQL = strSQL & " (JOBNO,SEQ,ENDDAY,ADJDAY,PURCHASENO,STD,ADJAMT,DIVAMT,DEMANDAMT,BALANCE,DIVFLAG,OUTSCODE,CUSER,CDATE)"
            strSQL = strSQL & " SELECT A.JOBNO,A.SEQ,'" & strENDDAY & "' ENDDAY,A.ADJDAY,A.PURCHASENO,A.STD,A.ADJAMT,B.DIVAMT,B.AMT DEMANDAMT,"
            strSQL = strSQL & " ISNULL(B.DIVAMT,0) - ISNULL(B.AMT,0) BALANCE,"
            strSQL = strSQL & " 'ACCRUED' DIVFLAG, "
            strSQL = strSQL & " A.OUTSCODE,'" & mobjSCGLConfig.WRKUSR & "' CUSER,'" & strNOW & "' CDATE"
            strSQL = strSQL & " FROM PD_EXE_DTL A LEFT JOIN "
            strSQL = strSQL & " ("
            strSQL = strSQL & " 	SELECT X.JOBNO,ISNULL(X.DIVAMT,0) DIVAMT,ISNULL(Y.AMT,0) AMT"
            strSQL = strSQL & " 	FROM"
            strSQL = strSQL & " 	(SELECT JOBNO,SUM(DIVAMT) DIVAMT FROM PD_DIVAMT GROUP BY JOBNO) X"
            strSQL = strSQL & " 	LEFT JOIN (SELECT JOBNO,SUM(AMT) AMT FROM PD_TAX_MST GROUP BY JOBNO) Y ON X.JOBNO = Y.JOBNO"
            strSQL = strSQL & " ) B ON A.JOBNO = B.JOBNO "
            strSQL = strSQL & " WHERE A.JOBNO = B.JOBNO "
            strSQL = strSQL & " AND A.JOBNO = '" & strJOBNO & "'"
            strSQL = strSQL & " AND ISNULL(A.PURCHASENO,'') = '' AND ISNULL(B.DIVAMT,0) - ISNULL(B.AMT,0) = 0"
            '추후 ISNULL(B.DIVAMT,0) - ISNULL(B.ADJAMT,0) = 0 부분은 사용자가 선택한 값으로 변경 될수 있음 견적과 청구가 불완전 해도 마감으로 할수 있다고 하면,,,,
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertClosing")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 마감처리시 결산데이터 생성
    '기타 : 선급금:PREPAYMENT 및 선청구:PREDEMANDAMT, 미지급금:ACCRUED 및 완료된내역:NULL  모두 저장
    '*****************************************************************
    Public Function InsertPREPAYMENT(ByVal strYEARMON As String) As Integer
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = " INSERT INTO PD_CLOSING_MST "
            strSQL = strSQL & " (YEARMON,JOBNO,SEQ,ENDDAY,ADJDAY,PURCHASENO,STD,ADJAMT,DIVAMT,DEMANDAMT,BALANCE,DIVFLAG,OUTSCODE,CUSER,CDATE)"
            strSQL = strSQL & " SELECT '" & strYEARMON & "' YEARMON, A.JOBNO,A.SEQ,C.ENDDAY,A.ADJDAY,A.PURCHASENO,A.STD,A.ADJAMT,B.DIVAMT,B.AMT DEMANDAMT,"
            strSQL = strSQL & " ISNULL(B.DIVAMT,0) - ISNULL(B.AMT,0) BALANCE,"
            strSQL = strSQL & " CASE WHEN ISNULL(A.PURCHASENO,'') <> '' AND ISNULL(B.DIVAMT,0) - ISNULL(B.AMT,0) <> 0 AND ISNULL(C.ENDDAY,'') = '' THEN 'PREPAYMENT'"
            strSQL = strSQL & " WHEN ISNULL(A.PURCHASENO,'') = '' AND ISNULL(B.DIVAMT,0) - ISNULL(B.AMT,0) = 0 AND ISNULL(C.ENDDAY,'') = '' THEN 'PREDEMANDAMT'"
            strSQL = strSQL & " WHEN ISNULL(A.PURCHASENO,'') = '' AND ISNULL(B.DIVAMT,0) - ISNULL(B.AMT,0) = 0 AND ISNULL(C.ENDDAY,'') <> '' THEN 'ACCRUED'"
            strSQL = strSQL & " END AS DIVFLAG,"
            strSQL = strSQL & " A.OUTSCODE,'" & mobjSCGLConfig.WRKUSR & "' CUSER,'" & strNOW & "' CDATE"
            strSQL = strSQL & " FROM PD_EXE_DTL A LEFT JOIN "
            strSQL = strSQL & " ("
            strSQL = strSQL & " 	SELECT X.JOBNO,ISNULL(X.DIVAMT,0) DIVAMT,ISNULL(Y.AMT,0) AMT"
            strSQL = strSQL & " 	FROM"
            strSQL = strSQL & " 	(SELECT JOBNO,SUM(DIVAMT) DIVAMT FROM PD_DIVAMT GROUP BY JOBNO) X"
            strSQL = strSQL & " 	LEFT JOIN (SELECT JOBNO,SUM(AMT) AMT FROM PD_TAX_MST GROUP BY JOBNO) Y ON X.JOBNO = Y.JOBNO"
            strSQL = strSQL & " ) B ON A.JOBNO = B.JOBNO"
            strSQL = strSQL & " LEFT JOIN PD_EXE_HDR C ON A.JOBNO = C.JOBNO "
            strSQL = strSQL & " WHERE A.JOBNO = B.JOBNO "
            strSQL = strSQL & " AND SUBSTRING(C.ENDDAY,1,6) = '" & strYEARMON & "' OR ISNULL(C.ENDDAY,'') = '';"
            '추후 ISNULL(B.DIVAMT,0) - ISNULL(B.ADJAMT,0) = 0 부분은 사용자가 선택한 값으로 변경 될수 있음 견적과 청구가 불완전 해도 마감으로 할수 있다고 하면,,,,

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertPREPAYMENT")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 / 마감처리시 정산마감으로 JOB 상태 변경
    '기타 : 
    '*****************************************************************
    Public Function InsertPREPAYMENT_AFTER(ByVal strYEARMON As String) As Integer
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try

            strSQL = " UPDATE PD_JOBNO SET ENDFLAG = 'PF05' "
            strSQL = strSQL & " FROM"
            strSQL = strSQL & " ("
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & " YEARMON,JOBNO,ENDDAY"
            strSQL = strSQL & " FROM PD_CLOSING_MST"
            strSQL = strSQL & " WHERE ISNULL(ENDDAY,'') <> ''"
            strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'  "
            strSQL = strSQL & " GROUP BY YEARMON,JOBNO,ENDDAY "
            strSQL = strSQL & " ) B"
            strSQL = strSQL & " WHERE PD_JOBNO.JOBNO = B.JOBNO;"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertPREPAYMENT_AFTER")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : 정산마감 취소처리
    '*****************************************************************
    Public Function DeletePREPAYMENT(ByVal strYEARMON As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_CLOSING_MST WHERE substring(YEARMON,1,6) = '" & strYEARMON & "' "

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeletePREPAYMENT")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : 정산마감 취소처리 - JOB상태 결산으로 조정
    '*****************************************************************
    Public Function DeletePREPAYMENT_BEFORE(ByVal strYEARMON As String) As Integer
        Dim strSQL As String

        Try
            strSQL = " UPDATE PD_JOBNO SET ENDFLAG = 'PF04' "
            strSQL = strSQL & " FROM"
            strSQL = strSQL & " ("
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & " YEARMON,JOBNO,ENDDAY"
            strSQL = strSQL & " FROM PD_CLOSING_MST"
            strSQL = strSQL & " WHERE ISNULL(ENDDAY,'') <> ''"
            strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'  "
            strSQL = strSQL & " GROUP BY YEARMON,JOBNO,ENDDAY "
            strSQL = strSQL & " ) B"
            strSQL = strSQL & " WHERE PD_JOBNO.JOBNO = B.JOBNO"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeletePREPAYMENT_BEFORE")
        End Try
    End Function
#End Region

#Region "객체 생성/해제"
    '*****************************************************************
    '입력 : strInfoXML = 공통기본정보에 대한 XML
    'objSCGLSql = DB 처리 객체 인스턴싱 변수    '반환 : 없음
    '기능 : DB 처리를 위한 공통기본정보 설정
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "PD_EXE_HDR"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

'------->>엔티티 INSERT/UPDATE 샘플입니다. 반드시 자신의 환경에 맞추어서 변경하시기 바랍니다.
'=========================================================
'       'vntData Array를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.InsertDo( _
'                                       GetElement(vntData,"SEQNO", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ACCCUSTCODE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"DEPTCD", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR01", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR02", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR03", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR04", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR05", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR06", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR07", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR08", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR09", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR10", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"CDATE", intColCnt, intRow, NULL_DTM, true ), _
'                                       GetElement(vntData,"UUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"UDATE", intColCnt, intRow, NULL_DTM, true ) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.UpdateDo( _
'                                       GetElement(vntData,"SEQNO", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ACCCUSTCODE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"DEPTCD", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR01", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR02", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR03", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR04", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR05", intColCnt, intRow), _
'                                       GetElement(vntData,"ATTR06", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR07", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR08", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR09", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"ATTR10", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"CDATE", intColCnt, intRow, NULL_DTM, true ), _
'                                       GetElement(vntData,"UUSER", intColCnt, intRow), _
'                                       GetElement(vntData,"UDATE", intColCnt, intRow, NULL_DTM, true ) _
'                                       )
'        Return intRtn


'=========================================================
'       'XmlData 를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.InsertDo( _
'                                       XMLGetElement(xmlRoot,"SEQNO"), _
'                                       XMLGetElement(xmlRoot,"SEQNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"ACCCUSTCODE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"DEPTCD"), _
'                                       XMLGetElement(xmlRoot,"ATTR01"), _
'                                       XMLGetElement(xmlRoot,"ATTR02"), _
'                                       XMLGetElement(xmlRoot,"ATTR03"), _
'                                       XMLGetElement(xmlRoot,"ATTR04"), _
'                                       XMLGetElement(xmlRoot,"ATTR05"), _
'                                       XMLGetElement(xmlRoot,"ATTR06", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR07", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR08", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR09", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR10", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSER"), _
'                                       XMLGetElement(xmlRoot,"CDATE", NULL_DTM, true ), _
'                                       XMLGetElement(xmlRoot,"UUSER"), _
'                                       XMLGetElement(xmlRoot,"UDATE", NULL_DTM, true ) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceSC_JOBCUST.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"SEQNO"), _
'                                       XMLGetElement(xmlRoot,"SEQNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"ACCCUSTCODE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"DEPTCD"), _
'                                       XMLGetElement(xmlRoot,"ATTR01"), _
'                                       XMLGetElement(xmlRoot,"ATTR02"), _
'                                       XMLGetElement(xmlRoot,"ATTR03"), _
'                                       XMLGetElement(xmlRoot,"ATTR04"), _
'                                       XMLGetElement(xmlRoot,"ATTR05"), _
'                                       XMLGetElement(xmlRoot,"ATTR06", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR07", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR08", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR09", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"ATTR10", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSER"), _
'                                       XMLGetElement(xmlRoot,"CDATE", NULL_DTM, true ), _
'                                       XMLGetElement(xmlRoot,"UUSER"), _
'                                       XMLGetElement(xmlRoot,"UDATE", NULL_DTM, true ) _
'                                       )
'        Return intRtn

