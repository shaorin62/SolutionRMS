
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

Public Class cePD_JOBNO
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_JOBNO"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    'PROJECTNO,JOBNO,JOBNAME,REQDAY,DEPTCD,EMPNO,HOPEENDDAY,BUDGETAMT,
    'JOBGUBN,CREGUBN,JOBBASE,CREDEPTCD,CREEMPNO,BIGO
    Public Function InsertDo(ByVal strPROJECTNO As String, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strJOBNAME As String = OPTIONAL_STR, _
            Optional ByVal strREQDAY As String = OPTIONAL_STR, _
            Optional ByVal strDEPTCD As String = OPTIONAL_STR, _
            Optional ByVal strEMPNO As String = OPTIONAL_STR, _
            Optional ByVal strHOPEENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
            Optional ByVal dblBUDGETAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strJOBGUBN As String = OPTIONAL_STR, _
            Optional ByVal strCREPART As String = OPTIONAL_STR, _
            Optional ByVal strCREGUBN As String = OPTIONAL_STR, _
            Optional ByVal strJOBBASE As String = OPTIONAL_STR, _
            Optional ByVal strCREDEPTCD As String = OPTIONAL_STR, _
            Optional ByVal strCREEMPNO As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strBIGO As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "PROJECTNO", strPROJECTNO, strFields, strValues)
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "JOBNAME", strJOBNAME, strFields, strValues)
            BuildNameValues(",", "REQDAY", strREQDAY, strFields, strValues)
            BuildNameValues(",", "DEPTCD", strDEPTCD, strFields, strValues)
            BuildNameValues(",", "EMPNO", strEMPNO, strFields, strValues)
            BuildNameValues(",", "HOPEENDDAY", strHOPEENDDAY, strFields, strValues)
            BuildNameValues(",", "ENDFLAG", strENDFLAG, strFields, strValues)
            BuildNameValues(",", "BUDGETAMT", dblBUDGETAMT, strFields, strValues)
            BuildNameValues(",", "JOBGUBN", strJOBGUBN, strFields, strValues)
            BuildNameValues(",", "CREPART", strCREPART, strFields, strValues)
            BuildNameValues(",", "CREGUBN", strCREGUBN, strFields, strValues)
            BuildNameValues(",", "JOBBASE", strJOBBASE, strFields, strValues)
            BuildNameValues(",", "CREDEPTCD", strCREDEPTCD, strFields, strValues)
            BuildNameValues(",", "CREEMPNO", strCREEMPNO, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "BIGO", strBIGO, strFields, strValues)
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
    'PROJECTNO,JOBNO,JOBNAME,REQDAY,DEPTCD,EMPNO,HOPEENDDAY,BUDGETAMT,
    'JOBGUBN,CREGUBN,JOBBASE,CREDEPTCD,CREEMPNO,BIGO
    Public Function UpdateDo(Optional ByVal strPROJECTNO As String = OPTIONAL_STR, _
            Optional ByVal strJOBNO As String = OPTIONAL_STR, _
            Optional ByVal strJOBNAME As String = OPTIONAL_STR, _
            Optional ByVal strREQDAY As String = OPTIONAL_STR, _
            Optional ByVal strDEPTCD As String = OPTIONAL_STR, _
            Optional ByVal strEMPNO As String = OPTIONAL_STR, _
            Optional ByVal strHOPEENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
            Optional ByVal dblBUDGETAMT As Double = OPTIONAL_NUM, _
            Optional ByVal strJOBGUBN As String = OPTIONAL_STR, _
            Optional ByVal strCREPART As String = OPTIONAL_STR, _
            Optional ByVal strCREGUBN As String = OPTIONAL_STR, _
            Optional ByVal strJOBBASE As String = OPTIONAL_STR, _
            Optional ByVal strCREDEPTCD As String = OPTIONAL_STR, _
            Optional ByVal strCREEMPNO As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strBIGO As String = OPTIONAL_STR, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        'PROJECTNO,JOBNO,JOBNAME,REQDAY,DEPTCD,EMPNO,HOPEENDDAY,BUDGETAMT,
        'JOBGUBN,CREGUBN,JOBBASE,CREDEPTCD,CREEMPNO,BIGO
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("JOBNAME", strJOBNAME), _
                        GetFieldNameValue("REQDAY", strREQDAY), _
                        GetFieldNameValue("DEPTCD", strDEPTCD), _
                        GetFieldNameValue("EMPNO", strEMPNO), _
                        GetFieldNameValue("HOPEENDDAY", strHOPEENDDAY), _
                        GetFieldNameValue("ENDFLAG", strENDFLAG), _
                        GetFieldNameValue("BUDGETAMT", dblBUDGETAMT), _
                        GetFieldNameValue("JOBGUBN", strJOBGUBN), _
                        GetFieldNameValue("CREPART", strCREPART), _
                        GetFieldNameValue("CREGUBN", strCREGUBN), _
                        GetFieldNameValue("JOBBASE", strJOBBASE), _
                        GetFieldNameValue("CREDEPTCD", strCREDEPTCD), _
                        GetFieldNameValue("CREEMPNO", strCREEMPNO), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("BIGO", strBIGO), _
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
                        GetFieldNameValue("PROJECTNO", strPROJECTNO), GetFieldNameValue("JOBNO", strJOBNO)))

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
    Public Function DeleteDo(Optional ByVal strSEQNO As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_JOBNO WHERE JOBNO = '" & strSEQNO & "';DELETE FROM PD_JOBNODEPT WHERE JOBNO = '" & strSEQNO & "';DELETE FROM PD_ACTUALRATE WHERE JOBNO = '" & strSEQNO & "'"
           
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
    '기능 : strSQL 해당문을 그대로 처리 완료구분의 진행상태를 의뢰상태로 변경
    '*****************************************************************
    Public Function UpdateRtn_ENDFLAG(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_ENDFLAG")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 :
    '*****************************************************************
    Public Function Update_CHECK_POINT(Optional ByVal strJOBNO As String = OPTIONAL_STR, Optional ByVal strCHECK_POINT As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format(" UPDATE {0} SET {1} WHERE {2}", EntityName, GetFieldNameValue("CHECK_POINT", strCHECK_POINT), _
                     BuildFields("AND", _
                        GetFieldNameValue("JOBNO", strJOBNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_CHECK_POINT")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Select 처리
    '*****************************************************************
    Public Function DeleteUpdateDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
                                   Optional ByVal strTRANSNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try

            strSQL = " UPDATE C SET  "
            strSQL = strSQL & " C.ENDFLAG = 'PF02' "
            strSQL = strSQL & " FROM PD_DIVAMT A , PD_TRANS_DTL B, PD_JOBNO C "
            strSQL = strSQL & " WHERE B.TRANSYEARMON = A.TRANSYEARMON AND B.TRANSNO = A.TRANSNO AND A.JOBNO = C.JOBNO"
            strSQL = strSQL & " AND B.TRANSYEARMON = '" & strTRANSYEARMON & "' AND  B.TRANSNO = '" & strTRANSNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteUpdateDo")
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
        MyBase.EntityName = "PD_JOBNO"     'Entity Name 설정
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

