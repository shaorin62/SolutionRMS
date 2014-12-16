'Public Class cePD_CONTRACT_MST

'End Class
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

Public Class cePD_CONTRACT_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_CONTRACT_MST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************


    Public Function InsertDo(ByVal strCONTRACTNO As String, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
            Optional ByVal strLOCALAREA As String = OPTIONAL_STR, _
            Optional ByVal strDELIVERYDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblPREAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCEAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblDELIVERYGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblFAULTGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal strPAYMENTGBN As String = OPTIONAL_STR, _
            Optional ByVal strSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strTESTDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTMENT As String = OPTIONAL_STR, _
            Optional ByVal dblTESTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblLOSTDAY As Double = OPTIONAL_NUM, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCONFLAG As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMENT As String = OPTIONAL_STR, _
            Optional ByVal strAMTFLAG As String = OPTIONAL_STR, _
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
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now


        Try
            BuildNameValues(",", "CONTRACTNO", strCONTRACTNO, strFields, strValues)
            BuildNameValues(",", "OUTSCODE", strOUTSCODE, strFields, strValues)
            BuildNameValues(",", "CONTRACTNAME", strCONTRACTNAME, strFields, strValues)
            BuildNameValues(",", "LOCALAREA", strLOCALAREA, strFields, strValues)
            BuildNameValues(",", "DELIVERYDAY", strDELIVERYDAY, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "PRERATE", dblPRERATE, strFields, strValues)
            BuildNameValues(",", "PREAMT", dblPREAMT, strFields, strValues)
            BuildNameValues(",", "ENDRATE", dblENDRATE, strFields, strValues)
            BuildNameValues(",", "ENDAMT", dblENDAMT, strFields, strValues)
            BuildNameValues(",", "THISRATE", dblTHISRATE, strFields, strValues)
            BuildNameValues(",", "THISAMT", dblTHISAMT, strFields, strValues)
            BuildNameValues(",", "BALANCERATE", dblBALANCERATE, strFields, strValues)
            BuildNameValues(",", "BALANCEAMT", dblBALANCEAMT, strFields, strValues)
            BuildNameValues(",", "DELIVERYGUARANTY", dblDELIVERYGUARANTY, strFields, strValues)
            BuildNameValues(",", "FAULTGUARANTY", dblFAULTGUARANTY, strFields, strValues)
            BuildNameValues(",", "PAYMENTGBN", strPAYMENTGBN, strFields, strValues)
            BuildNameValues(",", "STDATE", strSTDATE, strFields, strValues)
            BuildNameValues(",", "EDDATE", strEDDATE, strFields, strValues)
            BuildNameValues(",", "CONTRACTDAY", strCONTRACTDAY, strFields, strValues)
            BuildNameValues(",", "MANAGER", strMANAGER, strFields, strValues)
            BuildNameValues(",", "TESTDAY", strTESTDAY, strFields, strValues)
            BuildNameValues(",", "TESTENDDAY", strTESTENDDAY, strFields, strValues)
            BuildNameValues(",", "TESTMENT", strTESTMENT, strFields, strValues)
            BuildNameValues(",", "TESTAMT", dblTESTAMT, strFields, strValues)
            BuildNameValues(",", "LOSTDAY", dblLOSTDAY, strFields, strValues)
            BuildNameValues(",", "CONFIRMFLAG", strCONFIRMFLAG, strFields, strValues)
            BuildNameValues(",", "CONFLAG", strCONFLAG, strFields, strValues)
            BuildNameValues(",", "DIVFLAG", strDIVFLAG, strFields, strValues)
            BuildNameValues(",", "COMENT", strCOMENT, strFields, strValues)
            BuildNameValues(",", "AMTFLAG", strAMTFLAG, strFields, strValues)
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
    'CONTRACTNO,CONTRACTNAME,LOCALAREA,strSTDATE,strEDDATE,AMT,strDELIVERYDAY,strTESTDAY,PAYMENTGBN,TESTMENT,COMENT,CONFIRMFLAG
    Public Function UpdateDo(ByVal strCONTRACTNO As String, _
            Optional ByVal strOUTSCODE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTNAME As String = OPTIONAL_STR, _
            Optional ByVal strLOCALAREA As String = OPTIONAL_STR, _
            Optional ByVal strDELIVERYDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblPRERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblPREAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblENDAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISRATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblTHISAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCERATE As Double = OPTIONAL_NUM, _
            Optional ByVal dblBALANCEAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblDELIVERYGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal dblFAULTGUARANTY As Double = OPTIONAL_NUM, _
            Optional ByVal strPAYMENTGBN As String = OPTIONAL_STR, _
            Optional ByVal strSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strCONTRACTDAY As String = OPTIONAL_STR, _
            Optional ByVal strMANAGER As String = OPTIONAL_STR, _
            Optional ByVal strTESTDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTENDDAY As String = OPTIONAL_STR, _
            Optional ByVal strTESTMENT As String = OPTIONAL_STR, _
            Optional ByVal dblTESTAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblLOSTDAY As Double = OPTIONAL_NUM, _
            Optional ByVal strCONFIRMFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCONFLAG As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMENT As String = OPTIONAL_STR, _
            Optional ByVal strAMTFLAG As String = OPTIONAL_STR, _
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
        Try

            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CONTRACTNAME", strCONTRACTNAME), _
                        GetFieldNameValue("LOCALAREA", strLOCALAREA), _
                        GetFieldNameValue("DELIVERYDAY", strDELIVERYDAY), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("PRERATE", dblPRERATE), _
                        GetFieldNameValue("PREAMT", dblPREAMT), _
                        GetFieldNameValue("ENDRATE", dblENDRATE), _
                        GetFieldNameValue("ENDAMT", dblENDAMT), _
                        GetFieldNameValue("THISRATE", dblTHISRATE), _
                        GetFieldNameValue("THISAMT", dblTHISAMT), _
                        GetFieldNameValue("BALANCERATE", dblBALANCERATE), _
                        GetFieldNameValue("BALANCEAMT", dblBALANCEAMT), _
                        GetFieldNameValue("DELIVERYGUARANTY", dblDELIVERYGUARANTY), _
                        GetFieldNameValue("FAULTGUARANTY", dblFAULTGUARANTY), _
                        GetFieldNameValue("PAYMENTGBN", strPAYMENTGBN), _
                        GetFieldNameValue("STDATE", strSTDATE), _
                        GetFieldNameValue("EDDATE", strEDDATE), _
                        GetFieldNameValue("CONTRACTDAY", strCONTRACTDAY), _
                        GetFieldNameValue("MANAGER", strMANAGER), _
                        GetFieldNameValue("TESTDAY", strTESTDAY), _
                        GetFieldNameValue("TESTENDDAY", strTESTENDDAY), _
                        GetFieldNameValue("TESTMENT", strTESTMENT), _
                        GetFieldNameValue("TESTAMT", dblTESTAMT), _
                        GetFieldNameValue("LOSTDAY", dblLOSTDAY), _
                        GetFieldNameValue("CONFIRMFLAG", strCONFIRMFLAG), _
                        GetFieldNameValue("CONFLAG", strCONFLAG), _
                        GetFieldNameValue("DIVFLAG", strDIVFLAG), _
                        GetFieldNameValue("COMENT", strCOMENT), _
                        GetFieldNameValue("AMTFLAG", strAMTFLAG), _
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
                        GetFieldNameValue("CONTRACTNO", strCONTRACTNO)))

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
    'dblSEQ, strJOBNO, strPREESTNO
    Public Function DeleteDo(ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM PD_CONTRACT_MST WHERE CONTRACTNO = '" & strCONTRACTNO & "';"
            strSQL = strSQL & " UPDATE PD_EXE_DTL SET CONTRACTNO = '' WHERE CONTRACTNO = '" & strCONTRACTNO & "'; "
            strSQL = strSQL & " UPDATE PD_CONTRACT_NEW SET CONTRACTNO = '' WHERE CONTRACTNO = '" & strCONTRACTNO & "'; "

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
    '기능 : 해당 Entity에 Select 처리
    '*****************************************************************
    Public Function DeleteUpdateDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
           Optional ByVal strWOCODE As String = OPTIONAL_STR, _
           Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
           Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
           Optional ByVal strPREESTNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_DIVAMT SET  ADJAMT =  ADJAMT - " & dblADJAMT & ",ATTR02 = '',ATTR06 = NULL,ATTR07 = NULL WHERE SEQ = " & dblSEQ & " AND JOBNO = '" & strWOCODE & "'"
            strSQL = strSQL & " AND CLIENTCODE = '" & strCUSTCODE & "' AND PREESTNO = '" & strPREESTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteUpdateDo")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 진행비 상세내역 저장시 합계 금액 업데이트 
    '*****************************************************************
    Public Function Update_EXEHDR(ByVal strJOBNO As String, _
                                  ByVal dblAMT As Double) As Integer
        Dim strSQL As String
        'PAYMENT,RATE,INCOM,ACCAMT
        Try
            strSQL = " UPDATE PD_EXE_HDR SET ACCAMT = " & dblAMT & ","
            strSQL = strSQL & " INCOM = DEMANDAMT - (PAYMENT + " & dblAMT & "), "
            strSQL = strSQL & " RATE = ROUND(((DEMANDAMT - (PAYMENT+" & dblAMT & "))/DEMANDAMT)*100,2)"
            strSQL = strSQL & " WHERE JOBNO = '" & strJOBNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ADJAMTUpdateDo")
        End Try

    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 정산금액 업데이트 처리
    '*****************************************************************
    Public Function ADJAMTUpdateDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
                                   Optional ByVal strWOCODE As String = OPTIONAL_STR, _
                                   Optional ByVal dblADJAMT As Double = OPTIONAL_NUM, _
                                   Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
                                   Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                                   Optional ByVal dblATTR07 As Double = OPTIONAL_NUM, _
                                   Optional ByVal strPREESTNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = " UPDATE PD_DIVAMT SET ADJAMT = isnull(ADJAMT,0)+'" & dblADJAMT & "', "
            strSQL = strSQL & " ATTR06 = '" & dblATTR06 & "', "
            strSQL = strSQL & " ATTR02 = '" & strATTR02 & "',"
            strSQL = strSQL & " ATTR07 = '" & dblATTR07 & "' "
            strSQL = strSQL & " WHERE SEQ = '" & dblSEQ & "' AND JOBNO = '" & strWOCODE & "' AND PREESTNO = '" & strPREESTNO & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ADJAMTUpdateDo")
        End Try
    End Function
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : strSQL 해당문을 그대로 처리 소트순번 저장시 일괄업데이트 
    '*****************************************************************
    Public Function Update_CONTRACT(ByVal strTEMPSQL As String) As Integer
        Dim strSQL As String
        Try
            strSQL = strTEMPSQL
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_CONTRACT")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 정산금액 업데이트 처리
    '*****************************************************************
    Public Function UpdateDo_PD_EXE_DTL(Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                                        Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
                                        Optional ByVal strAMTFLAG As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = " UPDATE PD_EXE_DTL SET AMTFLAG = '" & strAMTFLAG & "' "
            strSQL = strSQL & " WHERE 1=1 AND JOBNO = '" & strJOBNO & "' AND SEQ = '" & dblSEQ & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".ADJAMTUpdateDo")
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
        MyBase.EntityName = "PD_CONTRACT_MST"     'Entity Name 설정
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

