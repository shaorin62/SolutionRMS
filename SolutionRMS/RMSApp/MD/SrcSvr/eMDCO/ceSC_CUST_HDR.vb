'****************************************************************************************
'Generated By:KTY
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceSC_CUST_HDR.vb ( SC_CUST_HDR Entity 처리 Class)
'기      능 : SC_CUST_DTL Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-06 By KTY
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceSC_CUST_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceSC_CUST_HDR"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strHIGHCUSTCODE As String, _
            Optional ByVal strACCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strGREATNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strREGNOTYPE As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strJURENTNO As String = OPTIONAL_STR, _
            Optional ByVal strBANKCODE As String = OPTIONAL_STR, _
            Optional ByVal strBANKACCNO As String = OPTIONAL_STR, _
            Optional ByVal strDEPNAME As String = OPTIONAL_STR, _
            Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
            Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
            Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
            Optional ByVal strTEL As String = OPTIONAL_STR, _
            Optional ByVal strFAX As String = OPTIONAL_STR, _
            Optional ByVal strOLDCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strKOBACOCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strUSE_FLAG As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR10 As Double = OPTIONAL_NUM, _
            Optional ByVal strCUSER As String = OPTIONAL_STR, _
            Optional ByVal strCDATE As String = OPTIONAL_STR, _
            Optional ByVal strUUSER As String = OPTIONAL_STR, _
            Optional ByVal strUDATE As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "HIGHCUSTCODE", strHIGHCUSTCODE, strFields, strValues)
            BuildNameValues(",", "ACCUSTCODE", strACCUSTCODE, strFields, strValues)
            BuildNameValues(",", "GREATCODE", strGREATCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "COMPANYNAME", strCOMPANYNAME, strFields, strValues)
            BuildNameValues(",", "GREATNAME", strGREATNAME, strFields, strValues)
            BuildNameValues(",", "CUSTTYPE", strCUSTTYPE, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "CUSTOWNER", strCUSTOWNER, strFields, strValues)
            BuildNameValues(",", "REGNOTYPE", strREGNOTYPE, strFields, strValues)
            BuildNameValues(",", "BUSINO", strBUSINO, strFields, strValues)
            BuildNameValues(",", "JURENTNO", strJURENTNO, strFields, strValues)
            BuildNameValues(",", "BANKCODE", strBANKCODE, strFields, strValues)
            BuildNameValues(",", "BANKACCNO", strBANKACCNO, strFields, strValues)
            BuildNameValues(",", "DEPNAME", strDEPNAME, strFields, strValues)
            BuildNameValues(",", "BUSISTAT", strBUSISTAT, strFields, strValues)
            BuildNameValues(",", "BUSITYPE", strBUSITYPE, strFields, strValues)
            BuildNameValues(",", "ZIPCODE", strZIPCODE, strFields, strValues)
            BuildNameValues(",", "ADDRESS1", strADDRESS1, strFields, strValues)
            BuildNameValues(",", "ADDRESS2", strADDRESS2, strFields, strValues)
            BuildNameValues(",", "TEL", strTEL, strFields, strValues)
            BuildNameValues(",", "FAX", strFAX, strFields, strValues)
            BuildNameValues(",", "OLDCUSTCODE", strOLDCUSTCODE, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "DEMANDFLAG", strDEMANDFLAG, strFields, strValues)
            BuildNameValues(",", "KOBACOCUSTCODE", strKOBACOCUSTCODE, strFields, strValues)
            BuildNameValues(",", "USE_FLAG", strUSE_FLAG, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", strATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", strATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", strATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", strATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", strATTR10, strFields, strValues)
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
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertCLIENT(ByVal strHIGHCUSTCODE As String, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strGREATNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
            Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
            Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
            Optional ByVal strTEL As String = OPTIONAL_STR, _
            Optional ByVal strFAX As String = OPTIONAL_STR, _
            Optional ByVal strUSE_FLAG As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR10 As Double = OPTIONAL_NUM, _
            Optional ByVal strCUSER As String = OPTIONAL_STR, _
            Optional ByVal strCDATE As String = OPTIONAL_STR, _
            Optional ByVal strUUSER As String = OPTIONAL_STR, _
            Optional ByVal strUDATE As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "HIGHCUSTCODE", strHIGHCUSTCODE, strFields, strValues)
            BuildNameValues(",", "GREATCODE", strGREATCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "COMPANYNAME", strCOMPANYNAME, strFields, strValues)
            BuildNameValues(",", "GREATNAME", strGREATNAME, strFields, strValues)
            BuildNameValues(",", "CUSTTYPE", strCUSTTYPE, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "CUSTOWNER", strCUSTOWNER, strFields, strValues)
            BuildNameValues(",", "BUSINO", strBUSINO, strFields, strValues)
            BuildNameValues(",", "BUSISTAT", strBUSISTAT, strFields, strValues)
            BuildNameValues(",", "BUSITYPE", strBUSITYPE, strFields, strValues)
            BuildNameValues(",", "ZIPCODE", strZIPCODE, strFields, strValues)
            BuildNameValues(",", "ADDRESS1", strADDRESS1, strFields, strValues)
            BuildNameValues(",", "ADDRESS2", strADDRESS2, strFields, strValues)
            BuildNameValues(",", "TEL", strTEL, strFields, strValues)
            BuildNameValues(",", "FAX", strFAX, strFields, strValues)
            BuildNameValues(",", "USE_FLAG", strUSE_FLAG, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", strATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", strATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", strATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", strATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", strATTR10, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)


            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)


            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertCLIENT")
        End Try
    End Function


    Public Function InsertEXCLIENT(ByVal strHIGHCUSTCODE As String, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
            Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
            Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
            Optional ByVal strTEL As String = OPTIONAL_STR, _
            Optional ByVal strFAX As String = OPTIONAL_STR, _
            Optional ByVal strUSE_FLAG As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR10 As Double = OPTIONAL_NUM, _
            Optional ByVal strCUSER As String = OPTIONAL_STR, _
            Optional ByVal strCDATE As String = OPTIONAL_STR, _
            Optional ByVal strUUSER As String = OPTIONAL_STR, _
            Optional ByVal strUDATE As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "HIGHCUSTCODE", strHIGHCUSTCODE, strFields, strValues)
            BuildNameValues(",", "GREATCODE", strGREATCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "COMPANYNAME", strCOMPANYNAME, strFields, strValues)
            BuildNameValues(",", "CUSTTYPE", strCUSTTYPE, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "CUSTOWNER", strCUSTOWNER, strFields, strValues)
            BuildNameValues(",", "BUSINO", strBUSINO, strFields, strValues)
            BuildNameValues(",", "BUSISTAT", strBUSISTAT, strFields, strValues)
            BuildNameValues(",", "BUSITYPE", strBUSITYPE, strFields, strValues)
            BuildNameValues(",", "ZIPCODE", strZIPCODE, strFields, strValues)
            BuildNameValues(",", "ADDRESS1", strADDRESS1, strFields, strValues)
            BuildNameValues(",", "ADDRESS2", strADDRESS2, strFields, strValues)
            BuildNameValues(",", "TEL", strTEL, strFields, strValues)
            BuildNameValues(",", "FAX", strFAX, strFields, strValues)
            BuildNameValues(",", "USE_FLAG", strUSE_FLAG, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", strATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", strATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", strATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", strATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", strATTR10, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)


            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)


            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertCLIENT")
        End Try
    End Function

    Public Function InsertMPPCLIENT(ByVal strHIGHCUSTCODE As String, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
            Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
            Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
            Optional ByVal strTEL As String = OPTIONAL_STR, _
            Optional ByVal strFAX As String = OPTIONAL_STR, _
            Optional ByVal strUSE_FLAG As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR01 As String = OPTIONAL_STR, _
            Optional ByVal strATTR02 As String = OPTIONAL_STR, _
            Optional ByVal strATTR03 As String = OPTIONAL_STR, _
            Optional ByVal strATTR04 As String = OPTIONAL_STR, _
            Optional ByVal strATTR05 As String = OPTIONAL_STR, _
            Optional ByVal strATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR07 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR08 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR09 As Double = OPTIONAL_NUM, _
            Optional ByVal strATTR10 As Double = OPTIONAL_NUM, _
            Optional ByVal strCUSER As String = OPTIONAL_STR, _
            Optional ByVal strCDATE As String = OPTIONAL_STR, _
            Optional ByVal strUUSER As String = OPTIONAL_STR, _
            Optional ByVal strUDATE As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "HIGHCUSTCODE", strHIGHCUSTCODE, strFields, strValues)
            BuildNameValues(",", "GREATCODE", strGREATCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "COMPANYNAME", strCOMPANYNAME, strFields, strValues)
            BuildNameValues(",", "CUSTTYPE", strCUSTTYPE, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "CUSTOWNER", strCUSTOWNER, strFields, strValues)
            BuildNameValues(",", "BUSINO", strBUSINO, strFields, strValues)
            BuildNameValues(",", "BUSISTAT", strBUSISTAT, strFields, strValues)
            BuildNameValues(",", "BUSITYPE", strBUSITYPE, strFields, strValues)
            BuildNameValues(",", "ZIPCODE", strZIPCODE, strFields, strValues)
            BuildNameValues(",", "ADDRESS1", strADDRESS1, strFields, strValues)
            BuildNameValues(",", "ADDRESS2", strADDRESS2, strFields, strValues)
            BuildNameValues(",", "TEL", strTEL, strFields, strValues)
            BuildNameValues(",", "FAX", strFAX, strFields, strValues)
            BuildNameValues(",", "USE_FLAG", strUSE_FLAG, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
            BuildNameValues(",", "ATTR06", strATTR06, strFields, strValues)
            BuildNameValues(",", "ATTR07", strATTR07, strFields, strValues)
            BuildNameValues(",", "ATTR08", strATTR08, strFields, strValues)
            BuildNameValues(",", "ATTR09", strATTR09, strFields, strValues)
            BuildNameValues(",", "ATTR10", strATTR10, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)


            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)


            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertCLIENT")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateDo(Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strACCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strGREATCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strGREATNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strREGNOTYPE As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strJURENTNO As String = OPTIONAL_STR, _
            Optional ByVal strBANKCODE As String = OPTIONAL_STR, _
            Optional ByVal strBANKACCNO As String = OPTIONAL_STR, _
            Optional ByVal strDEPNAME As String = OPTIONAL_STR, _
            Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
            Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
            Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
            Optional ByVal strTEL As String = OPTIONAL_STR, _
            Optional ByVal strFAX As String = OPTIONAL_STR, _
            Optional ByVal strOLDCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strKOBACOCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strUSE_FLAG As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("ACCUSTCODE", strACCUSTCODE), _
                        GetFieldNameValue("GREATCODE", strGREATCODE), _
                        GetFieldNameValue("CUSTNAME", strCUSTNAME), _
                        GetFieldNameValue("COMPANYNAME", strCOMPANYNAME), _
                        GetFieldNameValue("GREATNAME", strGREATNAME), _
                        GetFieldNameValue("CUSTTYPE", strCUSTTYPE), _
                        GetFieldNameValue("MEDFLAG", strMEDFLAG), _
                        GetFieldNameValue("CUSTOWNER", strCUSTOWNER), _
                        GetFieldNameValue("REGNOTYPE", strREGNOTYPE), _
                        GetFieldNameValue("BUSINO", strBUSINO), _
                        GetFieldNameValue("JURENTNO", strJURENTNO), _
                        GetFieldNameValue("BANKCODE", strBANKCODE), _
                        GetFieldNameValue("BANKACCNO", strBANKACCNO), _
                        GetFieldNameValue("DEPNAME", strDEPNAME), _
                        GetFieldNameValue("BUSISTAT", strBUSISTAT), _
                        GetFieldNameValue("BUSITYPE", strBUSITYPE), _
                        GetFieldNameValue("ZIPCODE", strZIPCODE), _
                        GetFieldNameValue("ADDRESS1", strADDRESS1), _
                        GetFieldNameValue("ADDRESS2", strADDRESS2), _
                        GetFieldNameValue("TEL", strTEL), _
                        GetFieldNameValue("FAX", strFAX), _
                        GetFieldNameValue("OLDCUSTCODE", strOLDCUSTCODE), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("DEMANDFLAG", strDEMANDFLAG), _
                        GetFieldNameValue("KOBACOCUSTCODE", strKOBACOCUSTCODE), _
                        GetFieldNameValue("USE_FLAG", strUSE_FLAG), _
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
                        GetFieldNameValue("HIGHCUSTCODE", strHIGHCUSTCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo1")
        End Try
    End Function

    Public Function UpdateCLIENT(Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
                                 Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
                                 Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
                                 Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
                                 Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
                                 Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
                                 Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
                                 Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
                                 Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
                                 Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
                                 Optional ByVal strTEL As String = OPTIONAL_STR, _
                                 Optional ByVal strFAX As String = OPTIONAL_STR, _
                                 Optional ByVal strUSE_FLAG As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("CUSTNAME", strCUSTNAME), _
                        GetFieldNameValue("COMPANYNAME", strCOMPANYNAME), _
                        GetFieldNameValue("CUSTTYPE", strCUSTTYPE), _
                        GetFieldNameValue("CUSTOWNER", strCUSTOWNER), _
                        GetFieldNameValue("BUSISTAT", strBUSISTAT), _
                        GetFieldNameValue("BUSITYPE", strBUSITYPE), _
                        GetFieldNameValue("ZIPCODE", strZIPCODE), _
                        GetFieldNameValue("ADDRESS1", strADDRESS1), _
                        GetFieldNameValue("ADDRESS2", strADDRESS2), _
                        GetFieldNameValue("TEL", strTEL), _
                        GetFieldNameValue("FAX", strFAX), _
                        GetFieldNameValue("USE_FLAG", strUSE_FLAG), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("HIGHCUSTCODE", strHIGHCUSTCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateCLIENT")
        End Try
    End Function


    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("HIGHCUSTCODE", strHIGHCUSTCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTY
    '반환 : 처리건수
    '기능 : 상위거래처가 미사용으로 변하면 하위거래처 일괄 미사용으로 변경
    '*****************************************************************
    Public Function Update_USEFLAG_DTL(ByVal strHIGHCUSTCODE As String, _
                                       ByVal strMEDFLAG As String) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = "UPDATE SC_CUST_DTL SET USE_FLAG = '0' "
            strSQL = strSQL & "  WHERE HIGHCUSTCODE = '" & strHIGHCUSTCODE & "' AND MEDFLAG = '" & strMEDFLAG & "'"
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_USEFLAG_DTL")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTY
    '반환 : 처리건수
    '기능 : 코바코 광고주코드를 KOBACOCUSTCODE컬럼에 UPDATE
    '*****************************************************************
    Public Function UpdateRtn_KOBACOCUSTCODE(ByVal strKOBACOCUSTCODE As String, _
                                             ByVal strHIGHCUSTCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE SC_CUST_HDR SET KOBACOCUSTCODE = '" & strKOBACOCUSTCODE & "' WHERE HIGHCUSTCODE = '" & strHIGHCUSTCODE & "' "

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_KOBACOCUSTCODE")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문 By KTY
    '반환 : 처리건수
    '기능 : 코바코 광고주코드를 KOBACOCUSTCODE컬럼에 UPDATE
    '*****************************************************************
    Public Function UpdateRtn_SBSCUSTCODE(ByVal strSBSCUSTCODE As String, _
                                          ByVal strHIGHCUSTCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE SC_CUST_HDR SET SBSCUSTCODE = '" & strSBSCUSTCODE & "' WHERE HIGHCUSTCODE = '" & strHIGHCUSTCODE & "' "

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_SBSCUSTCODE")
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
        MyBase.EntityName = "SC_CUST_HDR"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class
