'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceSC_REALMEDCODE_MST.vb ( SC_REALMEDCODE_MST Entity 처리 Class)
'기      능 : SC_REALMEDCODE_MST Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-14 오전 11:09:29 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceSC_VOCHNO
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceSC_VOCHNO"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(Optional ByVal strIF_KEY As String = OPTIONAL_STR, _
                             Optional ByVal strYEAR As String = OPTIONAL_STR, _
                             Optional ByVal strMON As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strGROUPSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                             Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal strTAXNO As String = OPTIONAL_STR, _
                             Optional ByVal strERR_YN As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "IF_KEY", strIF_KEY, strFields, strValues)
            BuildNameValues(",", "YEAR", strYEAR, strFields, strValues)
            BuildNameValues(",", "MON", strMON, strFields, strValues)
            BuildNameValues(",", "VOCHSEQ", strVOCHSEQ, strFields, strValues)
            BuildNameValues(",", "GROUPSEQ", strGROUPSEQ, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", strTAXNO, strFields, strValues)
            BuildNameValues(",", "ERR_YN", strERR_YN, strFields, strValues)
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
    Public Function UpdateDo(Optional ByVal strIF_KEY As String = OPTIONAL_STR, _
                             Optional ByVal strYEAR As String = OPTIONAL_STR, _
                             Optional ByVal strMON As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strGROUPSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                             Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal strTAXNO As String = OPTIONAL_STR, _
                             Optional ByVal strERR_YN As String = OPTIONAL_STR) As Integer
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("IF_KEY", strIF_KEY), _
                        GetFieldNameValue("YEAR", strYEAR), _
                        GetFieldNameValue("MON", strMON), _
                        GetFieldNameValue("VOCHSEQ", strVOCHSEQ), _
                        GetFieldNameValue("GROUPSEQ", strGROUPSEQ), _
                        GetFieldNameValue("VOCHNO", strVOCHNO), _
                        GetFieldNameValue("MEDFLAG", strMEDFLAG), _
                        GetFieldNameValue("TAXYEARMON", strTAXYEARMON), _
                        GetFieldNameValue("TAXNO", strTAXNO), _
                        GetFieldNameValue("ERR_YN", strERR_YN), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CODE", strVOCHNO)))
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
    Public Function DeleteDo(Optional ByVal dblSEQ As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SEQ", dblSEQ)))

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
        MyBase.EntityName = "SC_VOCHNO"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region

End Class






