'Public Class ceSC_BM

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

Public Class ceSC_CCTR
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceSC_CCTR"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(ByVal strBMCODE As String, _
                             Optional ByVal strHIGHDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strCCTR As String = OPTIONAL_STR, _
                             Optional ByVal strBA As String = OPTIONAL_STR, _
                             Optional ByVal strFDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strUSE_YN As String = OPTIONAL_STR) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "BMCODE", strBMCODE, strFields, strValues)
            BuildNameValues(",", "HIGHDEPT_CD", strHIGHDEPT_CD, strFields, strValues)
            BuildNameValues(",", "DEPT_CD", strDEPT_CD, strFields, strValues)
            BuildNameValues(",", "CCTR", strCCTR, strFields, strValues)
            BuildNameValues(",", "BA", strBA, strFields, strValues)
            BuildNameValues(",", "FDATE", strFDATE, strFields, strValues)
            BuildNameValues(",", "TDATE", strTDATE, strFields, strValues)
            BuildNameValues(",", "USE_YN", strUSE_YN, strFields, strValues)
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
    Public Function UpdateDo(ByVal dblSEQ As Double, _
                             ByVal strBMCODE As String, _
                             Optional ByVal strHIGHDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strDEPT_CD As String = OPTIONAL_STR, _
                             Optional ByVal strCCTR As String = OPTIONAL_STR, _
                             Optional ByVal strBA As String = OPTIONAL_STR, _
                             Optional ByVal strFDATE As String = OPTIONAL_STR, _
                             Optional ByVal strTDATE As String = OPTIONAL_STR, _
                             Optional ByVal strUSE_YN As String = OPTIONAL_STR) As Integer


        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = ""

            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                   BuildFields(",", _
                       GetFieldNameValue("BMCODE", strBMCODE), _
                       GetFieldNameValue("HIGHDEPT_CD", strHIGHDEPT_CD), _
                       GetFieldNameValue("DEPT_CD", strDEPT_CD), _
                       GetFieldNameValue("CCTR", strCCTR), _
                       GetFieldNameValue("BA", strBA), _
                       GetFieldNameValue("FDATE", strFDATE), _
                       GetFieldNameValue("TDATE", strTDATE), _
                       GetFieldNameValue("USE_YN", strUSE_YN)), _
                   BuildFields("AND", _
                     GetFieldNameValue("SEQ", dblSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
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
        MyBase.EntityName = "SC_CCTR"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class