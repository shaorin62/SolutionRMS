'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceSC_CUST_TEMP.vb ( SC_CUST_TEMP Entity 처리 Class)
'기      능 : SC_CUST_TEMP Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007-10-15 오후 1:14:17 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceSC_CUST_TEMP
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceSC_CUST_TEMP"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strCUSTCODE As String, _
            Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strCUSTOWNER As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strBUSISTAT As String = OPTIONAL_STR, _
            Optional ByVal strBUSITYPE As String = OPTIONAL_STR, _
            Optional ByVal strTEL As String = OPTIONAL_STR, _
            Optional ByVal strFAX As String = OPTIONAL_STR, _
            Optional ByVal strZIPCODE As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS1 As String = OPTIONAL_STR, _
            Optional ByVal strADDRESS2 As String = OPTIONAL_STR, _
            Optional ByVal strJURENTNO As String = OPTIONAL_STR, _
            Optional ByVal strBANKCODE As String = OPTIONAL_STR, _
            Optional ByVal strDEPNAME As String = OPTIONAL_STR, _
            Optional ByVal strBANKACCNO As String = OPTIONAL_STR, _
            Optional ByVal strMEDDIV As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strMPP As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "CUSTCODE", strCUSTCODE, strFields, strValues)
            BuildNameValues(",", "HIGHCUSTCODE", strHIGHCUSTCODE, strFields, strValues)
            BuildNameValues(",", "CUSTNAME", strCUSTNAME, strFields, strValues)
            BuildNameValues(",", "CUSTTYPE", strCUSTTYPE, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "BUSINO", strBUSINO, strFields, strValues)
            BuildNameValues(",", "DIVFLAG", strDIVFLAG, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "CUSTOWNER", strCUSTOWNER, strFields, strValues)
            BuildNameValues(",", "COMPANYNAME", strCOMPANYNAME, strFields, strValues)
            BuildNameValues(",", "BUSISTAT", strBUSISTAT, strFields, strValues)
            BuildNameValues(",", "BUSITYPE", strBUSITYPE, strFields, strValues)
            BuildNameValues(",", "TEL", strTEL, strFields, strValues)
            BuildNameValues(",", "FAX", strFAX, strFields, strValues)
            BuildNameValues(",", "ZIPCODE", strZIPCODE, strFields, strValues)
            BuildNameValues(",", "ADDRESS1", strADDRESS1, strFields, strValues)
            BuildNameValues(",", "ADDRESS2", strADDRESS2, strFields, strValues)
            BuildNameValues(",", "JURENTNO", strJURENTNO, strFields, strValues)
            BuildNameValues(",", "BANKCODE", strBANKCODE, strFields, strValues)
            BuildNameValues(",", "DEPNAME", strDEPNAME, strFields, strValues)
            BuildNameValues(",", "BANKACCNO", strBANKACCNO, strFields, strValues)
            BuildNameValues(",", "MEDDIV", strMEDDIV, strFields, strValues)
            BuildNameValues(",", "DEMANDFLAG", strDEMANDFLAG, strFields, strValues)
            BuildNameValues(",", "MPP", strMPP, strFields, strValues)
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
    Public Function UpdateDo1(Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strCOMPANYNAME As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal dblATTR06 As Double = OPTIONAL_NUM, _
            Optional ByVal strMEDDIV As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strMPP As String = OPTIONAL_STR, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer
        'XMLGetElement(xmlRoot, "MADDIV"), _
        'strDEMANDFLAG, _
        'XMLGetElement(xmlRoot, "MPP")
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("HIGHCUSTCODE", strHIGHCUSTCODE), _
                        GetFieldNameValue("CUSTNAME", strCUSTNAME), _
                        GetFieldNameValue("CUSTTYPE", strCUSTTYPE), _
                        GetFieldNameValue("MEDFLAG", strMEDFLAG), _
                        GetFieldNameValue("COMPANYNAME", strCOMPANYNAME), _
                        GetFieldNameValue("DIVFLAG", strDIVFLAG), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("ATTR06", dblATTR06), _
                        GetFieldNameValue("MEDDIV", strMEDDIV), _
                        GetFieldNameValue("DEMANDFLAG", strDEMANDFLAG), _
                        GetFieldNameValue("MPP", strMPP), _
                        GetFieldNameValue("ATTR10", dblATTR10), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CUSTCODE", strCUSTCODE)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo1")
        End Try
    End Function


    Public Function UpdateDo(Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strHIGHCUSTCODE As String = OPTIONAL_STR, _
            Optional ByVal strCUSTNAME As String = OPTIONAL_STR, _
            Optional ByVal strCUSTTYPE As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strBUSINO As String = OPTIONAL_STR, _
            Optional ByVal strDIVFLAG As String = OPTIONAL_STR, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
            Optional ByVal strMEDDIV As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDFLAG As String = OPTIONAL_STR, _
            Optional ByVal strMPP As String = OPTIONAL_STR, _
            Optional ByVal dblATTR10 As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("HIGHCUSTCODE", strHIGHCUSTCODE), _
                        GetFieldNameValue("CUSTNAME", strCUSTNAME), _
                        GetFieldNameValue("CUSTTYPE", strCUSTTYPE), _
                        GetFieldNameValue("MEDFLAG", strMEDFLAG), _
                        GetFieldNameValue("BUSINO", strBUSINO), _
                        GetFieldNameValue("DIVFLAG", strDIVFLAG), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("MEDDIV", strMEDDIV), _
                        GetFieldNameValue("DEMANDFLAG", strDEMANDFLAG), _
                        GetFieldNameValue("MPP", strMPP), _
                        GetFieldNameValue("ATTR10", dblATTR10), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CUSTCODE", strCUSTCODE)))

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
    Public Function DeleteDo(Optional ByVal strCUSTCODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("CUSTCODE", strCUSTCODE)))

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
                            Optional ByVal strCUSTCODE As String = OPTIONAL_STR, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("CUSTCODE", strCUSTCODE))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
        End Try
    End Function
    Public Function UpdateRtn_ATTR04(ByVal strBCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE SC_CUST_TEMP SET ATTR04 = '' WHERE ATTR04 = '" & strBCODE & "' "

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_ATTR04")
        End Try
    End Function

    Public Function UpdateRtn_CUSTCODE(ByVal strBCODE As String, _
                                       ByVal strCUSTCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE SC_CUST_TEMP SET ATTR04 = '" & strBCODE & "' WHERE CUSTCODE = '" & strCUSTCODE & "' "

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_CUSTCODE")
        End Try
    End Function
    Public Function UpdateRtn_ATTR03(ByVal strACODE As String, _
                                   ByVal strCLIENTCODE As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "UPDATE SC_CUST_TEMP SET ATTR03 = '" & strACODE & "' WHERE CUSTCODE = '" & strCLIENTCODE & "' "

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateRtn_ATTR03")
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
        MyBase.EntityName = "SC_CUST_TEMP"     'Entity Name 설정
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
'        intRtn = mobjceSC_CUST_TEMP.InsertDo( _
'                                       GetElement(vntData,"CUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"HIGHCUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ACCUSTCODE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSTNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTTYPE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"COMPANYNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTOWNER", intColCnt, intRow), _
'                                       GetElement(vntData,"REGNOTYPE", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSINO", intColCnt, intRow), _
'                                       GetElement(vntData,"JURENTNO", intColCnt, intRow), _
'                                       GetElement(vntData,"BANKCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"BANKACCNO", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSISTAT", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSITYPE", intColCnt, intRow), _
'                                       GetElement(vntData,"ZIPCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ADDRESS1", intColCnt, intRow), _
'                                       GetElement(vntData,"ADDRESS2", intColCnt, intRow), _
'                                       GetElement(vntData,"TEL", intColCnt, intRow), _
'                                       GetElement(vntData,"FAX", intColCnt, intRow), _
'                                       GetElement(vntData,"DIVFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"OLDCUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEMO", intColCnt, intRow), _
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
'        intRtn = mobjceSC_CUST_TEMP.UpdateDo( _
'                                       GetElement(vntData,"CUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"HIGHCUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ACCUSTCODE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CUSTNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTTYPE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"COMPANYNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"CUSTOWNER", intColCnt, intRow), _
'                                       GetElement(vntData,"REGNOTYPE", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSINO", intColCnt, intRow), _
'                                       GetElement(vntData,"JURENTNO", intColCnt, intRow), _
'                                       GetElement(vntData,"BANKCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"BANKACCNO", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPNAME", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSISTAT", intColCnt, intRow), _
'                                       GetElement(vntData,"BUSITYPE", intColCnt, intRow), _
'                                       GetElement(vntData,"ZIPCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"ADDRESS1", intColCnt, intRow), _
'                                       GetElement(vntData,"ADDRESS2", intColCnt, intRow), _
'                                       GetElement(vntData,"TEL", intColCnt, intRow), _
'                                       GetElement(vntData,"FAX", intColCnt, intRow), _
'                                       GetElement(vntData,"DIVFLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"OLDCUSTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEMO", intColCnt, intRow), _
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
'        intRtn = mobjceSC_CUST_TEMP.InsertDo( _
'                                       XMLGetElement(xmlRoot,"CUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"HIGHCUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"ACCUSTCODE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSTNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTTYPE"), _
'                                       XMLGetElement(xmlRoot,"MEDFLAG"), _
'                                       XMLGetElement(xmlRoot,"COMPANYNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTOWNER"), _
'                                       XMLGetElement(xmlRoot,"REGNOTYPE"), _
'                                       XMLGetElement(xmlRoot,"BUSINO"), _
'                                       XMLGetElement(xmlRoot,"JURENTNO"), _
'                                       XMLGetElement(xmlRoot,"BANKCODE"), _
'                                       XMLGetElement(xmlRoot,"BANKACCNO"), _
'                                       XMLGetElement(xmlRoot,"DEPNAME"), _
'                                       XMLGetElement(xmlRoot,"BUSISTAT"), _
'                                       XMLGetElement(xmlRoot,"BUSITYPE"), _
'                                       XMLGetElement(xmlRoot,"ZIPCODE"), _
'                                       XMLGetElement(xmlRoot,"ADDRESS1"), _
'                                       XMLGetElement(xmlRoot,"ADDRESS2"), _
'                                       XMLGetElement(xmlRoot,"TEL"), _
'                                       XMLGetElement(xmlRoot,"FAX"), _
'                                       XMLGetElement(xmlRoot,"DIVFLAG"), _
'                                       XMLGetElement(xmlRoot,"OLDCUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEMO"), _
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
'        intRtn = mobjceSC_CUST_TEMP.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"CUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"HIGHCUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"ACCUSTCODE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CUSTNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTTYPE"), _
'                                       XMLGetElement(xmlRoot,"MEDFLAG"), _
'                                       XMLGetElement(xmlRoot,"COMPANYNAME"), _
'                                       XMLGetElement(xmlRoot,"CUSTOWNER"), _
'                                       XMLGetElement(xmlRoot,"REGNOTYPE"), _
'                                       XMLGetElement(xmlRoot,"BUSINO"), _
'                                       XMLGetElement(xmlRoot,"JURENTNO"), _
'                                       XMLGetElement(xmlRoot,"BANKCODE"), _
'                                       XMLGetElement(xmlRoot,"BANKACCNO"), _
'                                       XMLGetElement(xmlRoot,"DEPNAME"), _
'                                       XMLGetElement(xmlRoot,"BUSISTAT"), _
'                                       XMLGetElement(xmlRoot,"BUSITYPE"), _
'                                       XMLGetElement(xmlRoot,"ZIPCODE"), _
'                                       XMLGetElement(xmlRoot,"ADDRESS1"), _
'                                       XMLGetElement(xmlRoot,"ADDRESS2"), _
'                                       XMLGetElement(xmlRoot,"TEL"), _
'                                       XMLGetElement(xmlRoot,"FAX"), _
'                                       XMLGetElement(xmlRoot,"DIVFLAG"), _
'                                       XMLGetElement(xmlRoot,"OLDCUSTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEMO"), _
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


