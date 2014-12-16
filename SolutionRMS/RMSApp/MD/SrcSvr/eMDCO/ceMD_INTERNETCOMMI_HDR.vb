'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_INTERNETCOMMI_HDR.vb ( MD_INTERNETCOMMI_HDR Entity 처리 Class)
'기      능 : MD_INTERNETCOMMI_HDR Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007-12-26 오후 5:41:28 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_INTERNETCOMMI_HDR
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_INTERNETCOMMI_HDR"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strTRANSYEARMON As String, _
            ByVal dblTRANSNO As Double, _
            ByVal strCLIENTCODE As String, _
            ByVal strREAL_MED_CODE As String, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
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
            BuildNameValues(",", "TRANSYEARMON", strTRANSYEARMON, strFields, strValues)
            BuildNameValues(",", "TRANSNO", dblTRANSNO, strFields, strValues)
            BuildNameValues(",", "CLIENTCODE", strCLIENTCODE, strFields, strValues)
            BuildNameValues(",", "REAL_MED_CODE", strREAL_MED_CODE, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "PRINTDAY", strPRINTDAY, strFields, strValues)
            BuildNameValues(",", "AMT", dblAMT, strFields, strValues)
            BuildNameValues(",", "VAT", dblVAT, strFields, strValues)
            BuildNameValues(",", "MED_FLAG", strMED_FLAG, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
            Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM, _
            Optional ByVal strCLIENTCODE As String = OPTIONAL_STR, _
            Optional ByVal strMEDCODE As String = OPTIONAL_STR, _
            Optional ByVal strREAL_MED_CODE As String = OPTIONAL_STR, _
            Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
            Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
            Optional ByVal strPRINTDAY As String = OPTIONAL_STR, _
            Optional ByVal dblAMT As Double = OPTIONAL_NUM, _
            Optional ByVal dblVAT As Double = OPTIONAL_NUM, _
            Optional ByVal strMED_FLAG As String = OPTIONAL_STR, _
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
                        GetFieldNameValue("CLIENTCODE", strCLIENTCODE), _
                        GetFieldNameValue("MEDCODE", strMEDCODE), _
                        GetFieldNameValue("REAL_MED_CODE", strREAL_MED_CODE), _
                        GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                        GetFieldNameValue("DEMANDDAY", strDEMANDDAY), _
                        GetFieldNameValue("PRINTDAY", strPRINTDAY), _
                        GetFieldNameValue("AMT", dblAMT), _
                        GetFieldNameValue("VAT", dblVAT), _
                        GetFieldNameValue("MED_FLAG", strMED_FLAG), _
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
                        GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO)))

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
    Public Function DeleteDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, Optional ByVal dblTRANSNO As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("TRANSYEARMON", strTRANSYEARMON), GetFieldNameValue("TRANSNO", dblTRANSNO)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

	'*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 거래명세서 완료건에 대하여 부가세 수정
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************

    Public Function Update_VatHDR(ByVal strTRANSYEARMON, ByVal strTRANSNO, ByVal lngSUMVAT) As Integer
        Dim strSQL As String

        Try
            strSQL = " UPDATE MD_INTERNETCOMMI_HDR "
            strSQL = strSQL & " SET VAT = " & lngSUMVAT
            strSQL = strSQL & " WHERE TRANSYEARMON = '" & strTRANSYEARMON & "' "
            strSQL = strSQL & " AND TRANSNO = " & strTRANSNO

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Update_VatHDR")
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
        MyBase.EntityName = "MD_INTERNETCOMMI_HDR"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class
