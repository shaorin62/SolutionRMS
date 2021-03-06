'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_INTERNET_MEDIUM.vb ( MD_INTERNET_MEDIUM Entity 처리 Class)
'기      능 : MD_INTERNET_MEDIUM Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-19 오후 7:19:00 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_INTERNET_CAMPAIGN
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_INTERNET_CAMPAIGN"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strCAMPAIGN_CODE As String, _
            ByVal strCAMPAIGN_NAME As String, _
            ByVal strCLIENTCODE As String, _
            ByVal strTIMCODE As String, _
            ByVal strSUBSEQ As String, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strTBRDSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strTBRDEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal dblMCCOMMI_RATE As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "CAMPAIGN_CODE", strCAMPAIGN_CODE, strFields, strValues)
            BuildNameValues(",", "CAMPAIGN_NAME", strCAMPAIGN_NAME, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "SUBSEQ", strSUBSEQ, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "TBRDSTDATE", strTBRDSTDATE, strFields, strValues)
            BuildNameValues(",", "TBRDEDDATE", strTBRDEDDATE, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "MCCOMMI_RATE", dblMCCOMMI_RATE, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strCAMPAIGN_CODE As String = OPTIONAL_STR, _
            Optional ByVal strCAMPAIGN_NAME As String = OPTIONAL_STR, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
            Optional ByVal strSUBSEQ As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strTBRDSTDATE As String = OPTIONAL_STR, _
            Optional ByVal strTBRDEDDATE As String = OPTIONAL_STR, _
            Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal dblMCCOMMI_RATE As Double = OPTIONAL_NUM, _
            Optional ByVal strMEMO As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("CAMPAIGN_NAME", strCAMPAIGN_NAME), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("TIMCODE", strTIMCODE), _
                        GetFieldNameValue("SUBSEQ", strSUBSEQ), _
                        GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                        GetFieldNameValue("TBRDSTDATE", strTBRDSTDATE), _
                        GetFieldNameValue("TBRDEDDATE", strTBRDEDDATE), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("MCCOMMI_RATE", dblMCCOMMI_RATE), _
                        GetFieldNameValue("MEMO", strMEMO), _
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
                        GetFieldNameValue("CAMPAIGN_CODE", strCAMPAIGN_CODE)))

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
    Public Function DeleteDo(Optional ByVal strCAMPAIGN_CODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("CAMPAIGN_CODE", strCAMPAIGN_CODE)))

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
                            Optional ByVal strYEARMON As String = OPTIONAL_STR, _
Optional ByVal dblSEQ As Double = OPTIONAL_NUM, _
                            Optional ByVal strSelFields As String = "*", _
                            Optional ByVal intLimitRow As Integer = 0, _
                            Optional ByVal intSelMode As Integer = SELMODE.ARR, _
                            Optional ByVal blnBindingHeader As Boolean = False) As Object
        Dim strSQL As String
        Dim strKeyFields As String

        Try
            strKeyFields = BuildFields("AND", _
                                    GetFieldNameValue("YEARMON", strYEARMON), GetFieldNameValue("SEQ", dblSEQ))

            Return SelectDoExt(intRowCnt, intColCnt, strSelFields, strKeyFields, intLimitRow, intSelMode, blnBindingHeader)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".SeleteDo")
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
        MyBase.EntityName = "MD_INTERNET_CAMPAIGN"     'Entity Name 설정
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
'        intRtn = mobjceMD_INTERNET_MEDIUM.InsertDo( _
'                                       GetElement(vntData,"YEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"SUBSEQ", intColCnt, intRow), _
'                                       GetElement(vntData,"PROGRAM_NAME", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"TBRDSTDATE", intColCnt, intRow), _
'                                       GetElement(vntData,"TBRDEDDATE", intColCnt, intRow), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMI_RATE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMISSION", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"NOTE", intColCnt, intRow), _
'                                       GetElement(vntData,"SPONSOR", intColCnt, intRow), _
'                                       GetElement(vntData,"GFLAG", intColCnt, intRow), _
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
'        intRtn = mobjceMD_INTERNET_MEDIUM.UpdateDo( _
'                                       GetElement(vntData,"YEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"SEQ", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"CLIENTCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDCODE", intColCnt, intRow), _
'                                       GetElement(vntData,"REAL_MED_CODE", intColCnt, intRow), _
'                                       GetElement(vntData,"MED_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"SUBSEQ", intColCnt, intRow), _
'                                       GetElement(vntData,"PROGRAM_NAME", intColCnt, intRow), _
'                                       GetElement(vntData,"DEPT_CD", intColCnt, intRow), _
'                                       GetElement(vntData,"TBRDSTDATE", intColCnt, intRow), _
'                                       GetElement(vntData,"TBRDEDDATE", intColCnt, intRow), _
'                                       GetElement(vntData,"AMT", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMI_RATE", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"COMMISSION", intColCnt, intRow, NULL_NUM, true ), _
'                                       GetElement(vntData,"TRU_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_FLAG", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRU_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TRANS_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_TAX_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"COMMI_VOCH_NO", intColCnt, intRow), _
'                                       GetElement(vntData,"NOTE", intColCnt, intRow), _
'                                       GetElement(vntData,"SPONSOR", intColCnt, intRow), _
'                                       GetElement(vntData,"GFLAG", intColCnt, intRow), _
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
'        intRtn = mobjceMD_INTERNET_MEDIUM.InsertDo( _
'                                       XMLGetElement(xmlRoot,"YEARMON"), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
'                                       XMLGetElement(xmlRoot,"SUBSEQ"), _
'                                       XMLGetElement(xmlRoot,"PROGRAM_NAME"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"TBRDSTDATE"), _
'                                       XMLGetElement(xmlRoot,"TBRDEDDATE"), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMI_RATE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMISSION", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"TRU_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"NOTE"), _
'                                       XMLGetElement(xmlRoot,"SPONSOR"), _
'                                       XMLGetElement(xmlRoot,"GFLAG"), _
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
'        intRtn = mobjceMD_INTERNET_MEDIUM.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"YEARMON"), _
'                                       XMLGetElement(xmlRoot,"SEQ", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"CLIENTCODE"), _
'                                       XMLGetElement(xmlRoot,"MEDCODE"), _
'                                       XMLGetElement(xmlRoot,"REAL_MED_CODE"), _
'                                       XMLGetElement(xmlRoot,"MED_FLAG"), _
'                                       XMLGetElement(xmlRoot,"SUBSEQ"), _
'                                       XMLGetElement(xmlRoot,"PROGRAM_NAME"), _
'                                       XMLGetElement(xmlRoot,"DEPT_CD"), _
'                                       XMLGetElement(xmlRoot,"TBRDSTDATE"), _
'                                       XMLGetElement(xmlRoot,"TBRDEDDATE"), _
'                                       XMLGetElement(xmlRoot,"AMT", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMI_RATE", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"COMMISSION", NULL_NUM, true ), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_FLAG"), _
'                                       XMLGetElement(xmlRoot,"TRU_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"TRU_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TRANS_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_TAX_NO"), _
'                                       XMLGetElement(xmlRoot,"COMMI_VOCH_NO"), _
'                                       XMLGetElement(xmlRoot,"NOTE"), _
'                                       XMLGetElement(xmlRoot,"SPONSOR"), _
'                                       XMLGetElement(xmlRoot,"GFLAG"), _
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


