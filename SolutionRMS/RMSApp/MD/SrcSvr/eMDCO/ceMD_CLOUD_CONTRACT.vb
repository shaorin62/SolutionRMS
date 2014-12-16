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

Public Class ceMD_CLOUD_CONTRACT
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_CLOUD_CONTRACT"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strCONT_CODE As String, _
                             ByVal strGUBUN As String, _
                             ByVal strCLIENTCODE As String, _
                             Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
                             Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strTBRDSTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTBRDEDDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONT_NAME As String = OPTIONAL_STR, _
                             Optional ByVal strCONT_TYPE As String = OPTIONAL_STR, _
                             Optional ByVal dblTOTAL_AMT As Double = OPTIONAL_NUM, _
                             Optional ByVal dblTIM_RATE As Double = OPTIONAL_NUM, _
                             Optional ByVal dblEX_RATE As Double = OPTIONAL_NUM, _
                             Optional ByVal dblCGV_RATE As Double = OPTIONAL_NUM, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR)



        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            BuildNameValues(",", "CONT_CODE", strCONT_CODE, strFields, strValues)
            BuildNameValues(",", "GUBUN", strGUBUN, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "TIMCODE", strTIMCODE, strFields, strValues)
            BuildNameValues(",", "EXCLIENTCODE", strEXCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "TBRDSTDATE", strTBRDSTDATE, strFields, strValues)
            BuildNameValues(",", "TBRDEDDATE", strTBRDEDDATE, strFields, strValues)
            BuildNameValues(",", "CONT_NAME", strCONT_NAME, strFields, strValues)
            BuildNameValues(",", "CONT_TYPE", strCONT_TYPE, strFields, strValues)
            BuildNameValues(",", "TOTAL_AMT", dblTOTAL_AMT, strFields, strValues)
            BuildNameValues(",", "TIM_RATE", dblTIM_RATE, strFields, strValues)
            BuildNameValues(",", "EX_RATE", dblEX_RATE, strFields, strValues)
            BuildNameValues(",", "CGV_RATE", dblCGV_RATE, strFields, strValues)
            BuildNameValues(",", "MEMO", strMEMO, strFields, strValues)
            BuildNameValues(",", "ENDFLAG", strENDFLAG, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "ATTR03", strATTR03, strFields, strValues)
            BuildNameValues(",", "ATTR04", strATTR04, strFields, strValues)
            BuildNameValues(",", "ATTR05", strATTR05, strFields, strValues)
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
    Public Function UpdateDo(ByVal strCONT_CODE As String, _
                             ByVal strGUBUN As String, _
                             ByVal strCLIENTCODE As String, _
                             Optional ByVal strTIMCODE As String = OPTIONAL_STR, _
                             Optional ByVal strEXCLIENTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strTBRDSTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTBRDEDDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCONT_NAME As String = OPTIONAL_STR, _
                             Optional ByVal strCONT_TYPE As String = OPTIONAL_STR, _
                             Optional ByVal dblTOTAL_AMT As Double = OPTIONAL_NUM, _
                             Optional ByVal dblTIM_RATE As Double = OPTIONAL_NUM, _
                             Optional ByVal dblEX_RATE As Double = OPTIONAL_NUM, _
                             Optional ByVal dblCGV_RATE As Double = OPTIONAL_NUM, _
                             Optional ByVal strMEMO As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR03 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR04 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR05 As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("GUBUN", strGUBUN), _
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("TIMCODE", strTIMCODE), _
                        GetFieldNameValue("EXCLIENTCODE", strEXCLIENTCODE), _
                        GetFieldNameValue("TBRDSTDATE", strTBRDSTDATE), _
                        GetFieldNameValue("TBRDEDDATE", strTBRDEDDATE), _
                        GetFieldNameValue("CONT_NAME", strCONT_NAME), _
                        GetFieldNameValue("CONT_TYPE", strCONT_TYPE), _
                        GetFieldNameValue("TOTAL_AMT", dblTOTAL_AMT), _
                        GetFieldNameValue("TIM_RATE", dblTIM_RATE), _
                        GetFieldNameValue("EX_RATE", dblEX_RATE), _
                        GetFieldNameValue("CGV_RATE", dblCGV_RATE), _
                        GetFieldNameValue("MEMO", strMEMO), _
                        GetFieldNameValue("ATTR01", strATTR01), _
                        GetFieldNameValue("ATTR02", strATTR02), _
                        GetFieldNameValue("ATTR03", strATTR03), _
                        GetFieldNameValue("ATTR04", strATTR04), _
                        GetFieldNameValue("ATTR05", strATTR05), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CONT_CODE", strCONT_CODE)))

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
    Public Function DeleteDo(Optional ByVal strCONT_CODE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("CONT_CODE", strCONT_CODE)))

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
        MyBase.EntityName = "MD_CLOUD_CONTRACT"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class