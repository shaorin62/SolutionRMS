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

Public Class cePD_CONTRACT_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_CONTRACT_HDR"    '자신의 클래스명
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
            Optional ByVal dblLOSTDAY As String = OPTIONAL_STR, _
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
    '참고 : Key 조건이 선택적임
    '*****************************************************************
    Public Function DeleteDo(ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = "DELETE FROM PD_CONTRACT_HDR WHERE CONTRACTNO = '" & strCONTRACTNO & "';"
            strSQL = strSQL & " UPDATE PD_CONTRACT_DTL "
            strSQL = strSQL & " SET CONTRACTNO = '', CONFIRM_USER = '" & mobjSCGLConfig.WRKUSR & "', CONFIRM_DATE = '" & strNOW & "'"
            strSQL = strSQL & " WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임
    '*****************************************************************
    Public Function DeleteRtn_Confirm_HDR(ByVal strCONTRACTNO As String) As Integer
        Dim strSQL As String
        Dim strNOW As String
        strNOW = Now

        Try
            strSQL = "DELETE FROM PD_CONTRACT_HDR WHERE CONTRACTNO = '" & strCONTRACTNO & "';"
            strSQL = strSQL & " UPDATE PD_CONTRACT_DTL "
            strSQL = strSQL & " SET CONTRACTNO = '', CONFIRM_USER = '" & mobjSCGLConfig.WRKUSR & "', CONFIRM_DATE = '" & strNOW & "'"
            strSQL = strSQL & " WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateDo_CONFIRM(ByVal strCONTRACTNO As String, _
                                     ByVal strCONFIRMFLAG As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE PD_CONTRACT_HDR SET CONFIRMFLAG = '" & strCONFIRMFLAG & "' WHERE CONTRACTNO = '" & strCONTRACTNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
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
        MyBase.EntityName = "PD_CONTRACT_HDR"     'Entity Name 설정
    End Sub

#End Region
#End Region
End Class
