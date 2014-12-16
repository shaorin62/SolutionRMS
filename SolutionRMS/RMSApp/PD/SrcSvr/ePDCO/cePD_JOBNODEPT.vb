
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

Public Class cePD_JOBNODEPT
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_JOBNODEPT"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"


#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    'SEQ|DEPTNAME|DEPTCODE|EMPNAME|EMPNO

    Public Function InsertDo(ByVal strJOBNO As String, _
                             Optional ByVal strSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strDEPTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strEMPNO As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNOSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strACTRATE As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "SEQ", strSEQ, strFields, strValues)
            BuildNameValues(",", "DEPTCODE", strDEPTCODE, strFields, strValues)
            BuildNameValues(",", "EMPNO", strEMPNO, strFields, strValues)
            BuildNameValues(",", "JOBNOSEQ", strJOBNOSEQ, strFields, strValues)
            BuildNameValues(",", "ACTRATE", strACTRATE, strFields, strValues)
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
    'strJOBNO,seq,deptcode,empno,jobnoseq,actrate
    Public Function UpdateDo(Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                             Optional ByVal strOLDSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strDEPTCODE As String = OPTIONAL_STR, _
                             Optional ByVal strEMPNO As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNOSEQ As Double = OPTIONAL_NUM, _
                             Optional ByVal strACTRATE As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("DEPTCODE", strDEPTCODE), _
                        GetFieldNameValue("EMPNO", strEMPNO), _
                        GetFieldNameValue("ACTRATE", strACTRATE), _
                        GetFieldNameValue("UUSER", mobjSCGLConfig.WRKUSR), _
                        GetFieldNameValue("UDATE", strNOW)), _
                    BuildFields("AND", _
                        GetFieldNameValue("SEQ", strOLDSEQ), GetFieldNameValue("JOBNO", strJOBNO), GetFieldNameValue("JOBNOSEQ", strJOBNOSEQ)))

            'BuildFields("AND", _
            '           GetFieldNameValue("SEQ", strOLDSEQ)), BuildFields("AND", _
            '                     GetFieldNameValue("JOBNO", strJOBNO)), BuildFields("AND", _
            '                    GetFieldNameValue("JOBNOSEQ", strJOBNOSEQ)))

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
    Public Function DeleteDo(Optional ByVal strSEQ As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                             Optional ByVal strJOBNOSEQ As Double = OPTIONAL_NUM) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("SEQ", strSEQ), _
                                   GetFieldNameValue("JOBNO", strJOBNO), _
                                   GetFieldNameValue("JOBNOSEQ", strJOBNOSEQ)))

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    Public Function DeleteDo2(Optional ByVal strJOBNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("JOBNO", strJOBNO)))

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
    Public Function DeleteDo_JOBNO(Optional ByVal strJOBNO As String = OPTIONAL_STR) As Integer

        Dim strSQL As String

        Try
            strSQL = String.Format("DELETE FROM {0} WHERE {1}", EntityName, _
                     BuildFields("AND", _
                                   GetFieldNameValue("JOBNO", strJOBNO)))

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
        MyBase.EntityName = "PD_JOBNODEPT"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

